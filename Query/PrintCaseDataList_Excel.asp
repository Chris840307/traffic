<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_建檔資料列表.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<%
Server.ScriptTimeout = 800
Response.flush
%>
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>建檔資料列表</title>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
'權限
'AuthorityCheck(234)

RecordDate=split(gInitDT(date),"-")

	strwhere=Session("PrintCarDataSQL")	
	Session.Contents.Remove("PrintCaseDataSQLxls")
	Session("PrintCaseDataSQLxls")=strwhere	

	'車輛
	strSQL="select * from BillBase a,MemberData b where a.RecordMemberID=b.MemberID(+)"&strwhere&" order by a.IllegalDate desc"
	set rsfound=conn.execute(strSQL)
	
	'行人
	strSQL2="select * from PasserBase a,MemberData b where a.RecordMemberID=b.MemberID(+)"&strwhere&" order by a.IllegalDate desc"
	set rsfound2=conn.execute(strSQL2)
	tmpSQL=strwhere

%>

</head>
<body>
<form name=myForm method="post">
<table width="100%" border="1" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
	<tr>
		<td bgcolor="#FFCC33" colspan="8">
			<font size="2"><strong>建檔資料清冊</strong></font>
		</td>
	</tr>
<%
i=0
If Not rsfound.Bof Then rsfound.MoveFirst 
While Not rsfound.Eof
		if i=0 then
			i=i+1
			TRcolor="#FFFFCC"
		else
			i=i-1
			TRcolor="#E1FFC4"
		end if
	'攔停
	if trim(rsfound("BillTypeID"))<>"2" then
%>
	<tr>
		<td align="right" bgcolor="<%=TRcolor%>" width="10%">單號</td>
		<td width="15%" align="left"><%
		if trim(rsfound("BillNo"))<>"" and not isnull(rsfound("BillNo")) then
			response.write trim(rsfound("BillNo"))
		else
			response.write "&nbsp;"
		end if
		%></td>

		<td align="right" bgcolor="<%=TRcolor%>" width="10%">違規種類</td>
		<td width="15%" align="left"><%
		if trim(rsfound("BillTypeID"))<>"" and not isnull(rsfound("BillTypeID")) then
			if trim(rsfound("BillTypeID"))="1" then
				response.write "攔停"
			elseif trim(rsfound("BillTypeID"))="2" then
				response.write "逕舉"
			end if
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>" width="10%">違規人姓名</td>
		<td width="15%" align="left"><%
		if trim(rsfound("Driver"))<>"" and not isnull(rsfound("Driver")) then
			response.write trim(rsfound("Driver"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>" width="10%">違規人證號</td>
		<td width="15%" align="left"><%
		if trim(rsfound("DriverID"))<>"" and not isnull(rsfound("DriverID")) then
			response.write trim(rsfound("DriverID"))
		else
			response.write "&nbsp;"
		end if
		%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="<%=TRcolor%>">違規人生日</td>
		<td align="left"><%
		if trim(rsfound("DriverBirth"))<>"" and not isnull(rsfound("DriverBirth")) then
			response.write gArrDT(trim(rsfound("DriverBirth")))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">違規人性別</td>
		<td align="left"><%
		if trim(rsfound("DriverSex"))<>"" and not isnull(rsfound("DriverSex")) then
			if trim(rsfound("DriverSex"))="1" then
				response.write "男"
			elseif trim(rsfound("DriverSex"))="2" then
				response.write "女"
			end if
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">違規人地址</td>
		<td colspan="3" align="left"><%
		if trim(rsfound("DriverZip"))<>"" and not isnull(rsfound("DriverZip")) then
				response.write trim(rsfound("DriverZip"))&"&nbsp;"
			end if
		if trim(rsfound("DriverAddress"))<>"" and not isnull(rsfound("DriverAddress")) then
				response.write trim(rsfound("DriverAddress"))
		else
			response.write "&nbsp;"
		end if
		%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="<%=TRcolor%>">車號</td>
		<td align="left"><%
		if trim(rsfound("CarNo"))<>"" and not isnull(rsfound("CarNo")) then
			response.write trim(rsfound("CarNo"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">簡示車種</td>
		<td align="left"><%
		if trim(rsfound("CarSimpleID"))<>"" and not isnull(rsfound("CarSimpleID")) then
			if trim(rsfound("CarSimpleID"))="1" then
				response.write "汽車"
			elseif trim(rsfound("CarSimpleID"))="2" then
				response.write "拖車"
			elseif trim(rsfound("CarSimpleID"))="3" then
				response.write "重機"
			elseif trim(rsfound("CarSimpleID"))="4" then
				response.write "輕機"
			end if
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">輔助車種</td>
		<td align="left"><%
		if trim(rsfound("CarAddID"))<>"" and not isnull(rsfound("CarAddID")) then
			if trim(rsfound("CarAddID"))="1" then
				response.write "大貨車"
			elseif trim(rsfound("CarAddID"))="2" then
				response.write "大客車"
			elseif trim(rsfound("CarAddID"))="3" then
				response.write "砂石車"
			elseif trim(rsfound("CarAddID"))="4" then
				response.write "土方車"
			elseif trim(rsfound("CarAddID"))="5" then
				response.write "動力機"
			elseif trim(rsfound("CarAddID"))="6" then
				response.write "貨櫃"
			end if
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">違規日期</td>
		<td align="left"><%
			if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
				response.write gArrDT(trim(rsfound("IllegalDate")))&"&nbsp; &nbsp;"
				response.write Right("00"&hour(rsfound("IllegalDate")),2)&":"
				response.write Right("00"&minute(rsfound("IllegalDate")),2)
			else
				response.write "&nbsp;"
			end if
		%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="<%=TRcolor%>">違規地點</td>
		<td colspan="3" align="left"><%
			if trim(rsfound("IllegalAddressID"))<>"" and not isnull(rsfound("IllegalAddressID")) then
				response.write trim(rsfound("IllegalAddressID"))&" "
			end if
			if trim(rsfound("IllegalAddress"))<>"" and not isnull(rsfound("IllegalAddress")) then
				response.write trim(rsfound("IllegalAddress"))
			else
				response.write "&nbsp;"
			end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">違規法條</td>
		<td align="left"><%
			if trim(rsfound("Rule1"))<>"" and not isnull(rsfound("Rule1")) then
				chRule=rsfound("Rule1")
			end if
			if trim(rsfound("Rule2"))<>"" and not isnull(rsfound("Rule2")) then
				chRule=chRule&"/"&rsfound("Rule2")
			end if
			if trim(rsfound("Rule3"))<>"" and not isnull(rsfound("Rule3")) then
				chRule=chRule&"/"&rsfound("Rule3")
			end if
			if trim(rsfound("Rule4"))<>"" and not isnull(rsfound("Rule4")) then 
				chRule=chRule&"/"&rsfound("Rule4")
			end if
			response.write chRule
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">限速、限重</td>
		<td align="left"><%
			if trim(rsfound("RuleSpeed"))<>"" and not isnull(rsfound("RuleSpeed")) then
				response.write trim(rsfound("RuleSpeed"))
			else
				response.write "&nbsp;"
			end if
		%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="<%=TRcolor%>">車速</td>
		<td align="left"><%
			if trim(rsfound("IllegalSpeed"))<>"" and not isnull(rsfound("IllegalSpeed")) then
				response.write trim(rsfound("IllegalSpeed"))
			else
				response.write "&nbsp;"
			end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">填單日期</td>
		<td align="left"><%
			if trim(rsfound("BillFillDate"))<>"" and not isnull(rsfound("BillFillDate")) then
				response.write gArrDT(trim(rsfound("BillFillDate")))
			else
				response.write "&nbsp;"
			end if
			%></td>
		<td align="right" bgcolor="<%=TRcolor%>">應到案日期</td>
		<td align="left"><%
			if trim(rsfound("DealLineDate"))<>"" and not isnull(rsfound("DealLineDate")) then
				response.write gArrDT(trim(rsfound("DealLineDate")))
			else
				response.write "&nbsp;"
			end if
			%></td>
		<td align="right" bgcolor="<%=TRcolor%>">應到案處所</td>
		<td align="left"><%
			if trim(rsfound("MemberStation"))<>"" and not isnull(rsfound("MemberStation")) then
				strMStation="select stationName from Station where StationID='"&trim(rsfound("MemberStation"))&"'"
				set rsMStation=conn.execute(strMStation)
				if not rsMStation.eof then
					response.write trim(rsMStation("stationName"))
				end if
				rsMStation.close
				set rsMStation=nothing
			else
				response.write "&nbsp;"
			end if
		%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="<%=TRcolor%>">舉發單位</td>
		<td align="left"><%
			if trim(rsfound("BillUnitID"))<>"" and not isnull(rsfound("BillUnitID")) then
				strUName="select UnitName from UnitInfo where UnitID='"&trim(rsfound("BillUnitID"))&"'"
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
		<td align="right" bgcolor="<%=TRcolor%>">舉發人員</td>
		<td align="left"><%
			if trim(rsfound("BillMem1"))<>"" and not isnull(rsfound("BillMem1")) then
				response.write trim(rsfound("BillMem1"))
			else
				response.write "&nbsp;"
			end if
			if trim(rsfound("BillMem2"))<>"" and not isnull(rsfound("BillMem2")) then
				response.write "、"&trim(rsfound("BillMem2"))
			end if
			if trim(rsfound("BillMem3"))<>"" and not isnull(rsfound("BillMem3")) then
				response.write "、"&trim(rsfound("BillMem3"))
			end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">代保管物</td>
		<td align="left"><%
			FastenerDetail=""
			strFas="select b.Content from BillFastenerDetail a,DCICode b where BillSN="&trim(rsfound("SN"))&" and b.TypeID=6 and a.FastenerTypeID=b.ID"
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
		<td align="right" bgcolor="<%=TRcolor%>">專案</td>
		<td align="left"><%
			if trim(rsfound("ProjectID"))<>"" and not isnull(rsfound("ProjectID")) then
				strProj="select Name from Project where ProjectID='"&trim(rsfound("ProjectID"))&"'"
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
		<td align="right" bgcolor="<%=TRcolor%>">第三責任險</td>
		<td align="left"><%
			if trim(rsfound("Insurance"))<>"" and not isnull(rsfound("Insurance")) then
				if trim(rsfound("Insurance"))="0" then
					response.write "有出示"
				elseif trim(rsfound("Insurance"))="1" then
					response.write "未出示"
				elseif trim(rsfound("Insurance"))="2" then
					response.write "肇事且未出示"
				elseif trim(rsfound("Insurance"))="3" then
					response.write "逾期或未保險"
				elseif trim(rsfound("Insurance"))="4" then
					response.write "肇事且逾期或未保險"
				end if
			else
				response.write "&nbsp;"
			end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">交通事故案號</td>
		<td align="left"><%
			if trim(rsfound("TrafficAccidentNo"))<>"" and not isnull(rsfound("TrafficAccidentNo")) then
				response.write trim(rsfound("TrafficAccidentNo"))
			else
				response.write "&nbsp;"
			end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">交通事故種類</td>
		<td align="left"><%
			if trim(rsfound("TrafficAccidentType"))<>"" and not isnull(rsfound("TrafficAccidentType")) then
				response.write trim(rsfound("TrafficAccidentType"))
			else
				response.write "&nbsp;"
			end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">備註</td>
		<td align="left"><%
			if trim(rsfound("Note"))<>"" and not isnull(rsfound("Note")) then
				response.write trim(rsfound("Note"))
			else
				response.write "&nbsp;"
			end if
		%></td>
	</tr>
<%	'逕舉
	elseif trim(rsfound("BillTypeID"))="2" then
%>
	<tr>
		<td align="right" bgcolor="<%=TRcolor%>" width="10%">單號</td>
		<td width="15%" align="left"><%
		if trim(rsfound("BillNo"))<>"" and not isnull(rsfound("BillNo")) then
			response.write trim(rsfound("BillNo"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>" width="10%">違規種類</td>
		<td width="15%" align="left"><%
		if trim(rsfound("BillTypeID"))<>"" and not isnull(rsfound("BillTypeID")) then
			if trim(rsfound("BillTypeID"))="1" then
				response.write "攔停"
			elseif trim(rsfound("BillTypeID"))="2" then
				response.write "逕舉"
			end if
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>" width="10%">採証工具</td>
		<td width="15%" align="left"><%
		'(空:無/1:固定桿/2:雷達測速[三腳架]/3:儀器[相機])
			if trim(rsfound("UseTool"))<>"" and not isnull(rsfound("UseTool")) then
				if trim(rsfound("UseTool"))="1" then
					response.write "固定桿"
				elseif trim(rsfound("UseTool"))="2" then
					response.write "雷達測速[三腳架]"
				elseif trim(rsfound("UseTool"))="3" then
					response.write "儀器[相機]"
				end if
			else
				response.write "無"
			end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>" width="10%">固定桿編號</td>
		<td width="15%" align="left"><%
		if trim(rsfound("EquipmentID"))<>"" and not isnull(rsfound("EquipmentID")) then
			response.write trim(rsfound("EquipmentID"))
		else
			response.write "&nbsp;"
		end if
		%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="<%=TRcolor%>">車號</td>
		<td width="15%" align="left"><%
		if trim(rsfound("CarNo"))<>"" and not isnull(rsfound("CarNo")) then
			response.write trim(rsfound("CarNo"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">簡示車種</td>
		<td align="left"><%
		if trim(rsfound("CarSimpleID"))<>"" and not isnull(rsfound("CarSimpleID")) then
			if trim(rsfound("CarSimpleID"))="1" then
				response.write "汽車"
			elseif trim(rsfound("CarSimpleID"))="2" then
				response.write "拖車"
			elseif trim(rsfound("CarSimpleID"))="3" then
				response.write "重機"
			elseif trim(rsfound("CarSimpleID"))="4" then
				response.write "輕機"
			else
				response.write "&nbsp;"
			end if
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">輔助車種</td>
		<td align="left"><%
		if trim(rsfound("CarAddID"))<>"" and not isnull(rsfound("CarAddID")) then
			if trim(rsfound("CarAddID"))="1" then
				response.write "大貨車"
			elseif trim(rsfound("CarAddID"))="2" then
				response.write "大客車"
			elseif trim(rsfound("CarAddID"))="3" then
				response.write "砂石車"
			elseif trim(rsfound("CarAddID"))="4" then
				response.write "土方車"
			elseif trim(rsfound("CarAddID"))="5" then
				response.write "動力機"
			elseif trim(rsfound("CarAddID"))="6" then
				response.write "貨櫃"
			else
				response.write "&nbsp;"
			end if
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">違規日期</td>
		<td align="left"><%
			if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
				response.write gArrDT(trim(rsfound("IllegalDate")))&"&nbsp;&nbsp;"
				response.write Right("00"&hour(rsfound("IllegalDate")),2)&":"
				response.write Right("00"&minute(rsfound("IllegalDate")),2)
			else
				response.write "&nbsp;"
			end if
		%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="<%=TRcolor%>">違規地點</td>
		<td colspan="3" align="left"><%
			if trim(rsfound("IllegalAddressID"))<>"" and not isnull(rsfound("IllegalAddressID")) then
				response.write trim(rsfound("IllegalAddressID"))&" "
			end if
			if trim(rsfound("IllegalAddress"))<>"" and not isnull(rsfound("IllegalAddress")) then
				response.write trim(rsfound("IllegalAddress"))
			else
				response.write "&nbsp;"
			end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">違規法條</td>
		<td align="left"><%
			if trim(rsfound("Rule1"))<>"" and not isnull(rsfound("Rule1")) then
				response.write rsfound("Rule1")
			else
				response.write "&nbsp;"
			end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">限速、限重</td>
		<td align="left"><%
			if trim(rsfound("RuleSpeed"))<>"" and not isnull(rsfound("RuleSpeed")) then
				response.write trim(rsfound("RuleSpeed"))
			else
				response.write "&nbsp;"
			end if
		%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="<%=TRcolor%>">車速</td>
		<td align="left"><%
			if trim(rsfound("IllegalSpeed"))<>"" and not isnull(rsfound("IllegalSpeed")) then
				response.write trim(rsfound("IllegalSpeed"))
			else
				response.write "&nbsp;"
			end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">填單日期</td>
		<td align="left"><%
			if trim(rsfound("BillFillDate"))<>"" and not isnull(rsfound("BillFillDate")) then
				response.write gArrDT(trim(rsfound("BillFillDate")))
			else
				response.write "&nbsp;"
			end if
			%></td>
		<td align="right" bgcolor="<%=TRcolor%>">應到案日期</td>
		<td align="left"><%
			if trim(rsfound("DealLineDate"))<>"" and not isnull(rsfound("DealLineDate")) then
				response.write gArrDT(trim(rsfound("DealLineDate")))
			else
				response.write "&nbsp;"
			end if
			%></td>
		<td align="right" bgcolor="<%=TRcolor%>">應到案處所</td>
		<td align="left"><%
			if trim(rsfound("MemberStation"))<>"" and not isnull(rsfound("MemberStation")) then
				strMStation="select stationName from Station where StationID='"&trim(rsfound("MemberStation"))&"'"
				set rsMStation=conn.execute(strMStation)
				if not rsMStation.eof then
					response.write trim(rsMStation("stationName"))
				end if
				rsMStation.close
				set rsMStation=nothing
			else
				response.write "&nbsp;"
			end if
		%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="<%=TRcolor%>">舉發單位</td>
		<td align="left"><%
			if trim(rsfound("BillUnitID"))<>"" and not isnull(rsfound("BillUnitID")) then
				strUName="select UnitName from UnitInfo where UnitID='"&trim(rsfound("BillUnitID"))&"'"
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
		<td align="right" bgcolor="<%=TRcolor%>">舉發人員</td>
		<td align="left"><%
			if trim(rsfound("BillMem1"))<>"" and not isnull(rsfound("BillMem1")) then
				response.write trim(rsfound("BillMem1"))
			else
				response.write "&nbsp;"
			end if
			if trim(rsfound("BillMem2"))<>"" and not isnull(rsfound("BillMem2")) then
				response.write "、"&trim(rsfound("BillMem2"))
			end if
			if trim(rsfound("BillMem3"))<>"" and not isnull(rsfound("BillMem3")) then
				response.write "、"&trim(rsfound("BillMem3"))
			end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">專案</td>
		<td align="left"><%
			if trim(rsfound("ProjectID"))<>"" and not isnull(rsfound("ProjectID")) then
				strProj="select Name from Project where ProjectID='"&trim(rsfound("ProjectID"))&"'"
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
		<td align="right" bgcolor="<%=TRcolor%>">備註</td>
		<td align="left"><%
			if trim(rsfound("Note"))<>"" and not isnull(rsfound("Note")) then
				response.write trim(rsfound("Note"))
			else
				response.write "&nbsp;"
			end if
		%></td>
	</tr>
<%
	end if
%>
	<tr>
		<td colspan="8"></td>
	</tr>
<%
rsfound.MoveNext
Wend
rsfound.close
set rsfound=nothing
%>
<%
j=i
If Not rsfound2.Bof Then rsfound2.MoveFirst 
While Not rsfound2.Eof
		if j=0 then
			j=j+1
			TRcolor="#FFFFCC"
		else
			j=j-1
			TRcolor="#E1FFC4"
		end if
%>
	<tr>
		<td align="right" bgcolor="<%=TRcolor%>" width="10%">單號</td>
		<td width="15%" align="left"><%
		if trim(rsfound2("BillNo"))<>"" and not isnull(rsfound2("BillNo")) then
			response.write trim(rsfound2("BillNo"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>" width="10%">告發類別</td>
		<td width="15%" align="left"><%
		if trim(rsfound2("BillTypeID"))<>"" and not isnull(rsfound2("BillTypeID")) then
			if trim(rsfound2("BillTypeID"))="1" then
				response.write "慢車"
			elseif trim(rsfound2("BillTypeID"))="2" then
				response.write "行人"
			elseif trim(rsfound2("BillTypeID"))="3" then
				response.write "道路障礙"
			end if
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>" width="10%">是否參加講習</td>
		<td width="15%" align="left"><%
		if trim(rsfound2("IsLecture"))<>"" and not isnull(rsfound2("IsLecture")) then
			if trim(rsfound2("IsLecture"))="0" then
				response.write "否"
			elseif trim(rsfound2("IsLecture"))="1" then
				response.write "是"
			end if
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>" width="10%">違規人姓名</td>
		<td width="15%" align="left"><%
		if trim(rsfound2("Driver"))<>"" and not isnull(rsfound2("Driver")) then
			response.write trim(rsfound2("Driver"))
		else
			response.write "&nbsp;"
		end if
		%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="<%=TRcolor%>" width="10%">違規人證號</td>
		<td width="15%" align="left"><%
		if trim(rsfound2("DriverID"))<>"" and not isnull(rsfound2("DriverID")) then
			response.write trim(rsfound2("DriverID"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">違規人生日</td>
		<td align="left"><%
		if trim(rsfound2("DriverBirth"))<>"" and not isnull(rsfound2("DriverBirth")) then
			response.write gArrDT(trim(rsfound2("DriverBirth")))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">違規人地址</td>
		<td colspan="3" align="left"><%
		if trim(rsfound2("DriverZip"))<>"" and not isnull(rsfound2("DriverZip")) then
				response.write trim(rsfound2("DriverZip"))&"&nbsp;"
			end if
		if trim(rsfound2("DriverAddress"))<>"" and not isnull(rsfound2("DriverAddress")) then
				response.write trim(rsfound2("DriverAddress"))
		else
			response.write "&nbsp;"
		end if
		%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="<%=TRcolor%>">違規日期</td>
		<td align="left"><%
			if trim(rsfound2("IllegalDate"))<>"" and not isnull(rsfound2("IllegalDate")) then
				response.write gArrDT(trim(rsfound2("IllegalDate")))&"&nbsp;&nbsp;"
				response.write Right("00"&hour(rsfound2("IllegalDate")),2)&":"
				response.write Right("00"&minute(rsfound2("IllegalDate")),2)
			else
				response.write "&nbsp;"
			end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">違規地點</td>
		<td colspan="3" align="left"><%
			if trim(rsfound2("IllegalAddressID"))<>"" and not isnull(rsfound2("IllegalAddressID")) then
				response.write trim(rsfound2("IllegalAddressID"))&" "
			end if
			if trim(rsfound2("IllegalAddress"))<>"" and not isnull(rsfound2("IllegalAddress")) then
				response.write trim(rsfound2("IllegalAddress"))
			else
				response.write "&nbsp;"
			end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">違規法條</td>
		<td align="left"><%
			if trim(rsfound2("Rule1"))<>"" and not isnull(rsfound2("Rule1")) then
				response.write rsfound2("Rule1")
			else
				response.write "&nbsp;"
			end if
		%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="<%=TRcolor%>">填單日期</td>
		<td align="left"><%
			if trim(rsfound2("BillFillDate"))<>"" and not isnull(rsfound2("BillFillDate")) then
				response.write gArrDT(trim(rsfound2("BillFillDate")))
			else
				response.write "&nbsp;"
			end if
			%></td>
		<td align="right" bgcolor="<%=TRcolor%>">應到案日期</td>
		<td align="left"><%
			if trim(rsfound2("DealLineDate"))<>"" and not isnull(rsfound2("DealLineDate")) then
				response.write gArrDT(trim(rsfound2("DealLineDate")))
			else
				response.write "&nbsp;"
			end if
			%></td>
		<td align="right" bgcolor="<%=TRcolor%>">應到案處所</td>
		<td align="left"><%
			if trim(rsfound2("MemberStation"))<>"" and not isnull(rsfound2("MemberStation")) then
				strMStation="select UnitName from UnitInfo where UnitID='"&trim(rsfound2("MemberStation"))&"'"
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
		<td align="right" bgcolor="<%=TRcolor%>">舉發單位</td>
		<td align="left"><%
			if trim(rsfound2("BillUnitID"))<>"" and not isnull(rsfound2("BillUnitID")) then
				strUName="select UnitName from UnitInfo where UnitID='"&trim(rsfound2("BillUnitID"))&"'"
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
		<td align="right" bgcolor="<%=TRcolor%>">舉發人員</td>
		<td align="left"><%
			if trim(rsfound2("BillMem1"))<>"" and not isnull(rsfound2("BillMem1")) then
				response.write trim(rsfound2("BillMem1"))
			else
				response.write "&nbsp;"
			end if
			if trim(rsfound2("BillMem2"))<>"" and not isnull(rsfound2("BillMem2")) then
				response.write "、"&trim(rsfound2("BillMem2"))
			end if
			if trim(rsfound2("BillMem3"))<>"" and not isnull(rsfound2("BillMem3")) then
				response.write "、"&trim(rsfound2("BillMem3"))
			end if
		%></td>
		<td align="right" bgcolor="<%=TRcolor%>">代保管物</td>
		<td align="left"><%
			FastenerDetail=""
			strFas="select Confiscate from PasserConfiscate where BillSN="&trim(rsfound2("SN"))
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
		<td align="right" bgcolor="<%=TRcolor%>">備註</td>
		<td colspan="3" align="left"><%
			if trim(rsfound2("Note"))<>"" and not isnull(rsfound2("Note")) then
				response.write trim(rsfound2("Note"))
			else
				response.write "&nbsp;"
			end if
		%></td>
	</tr>
	<tr>
		<td colspan="8"></td>
	</tr>
<%
rsfound2.MoveNext
Wend
rsfound2.close
set rsfound2=nothing
%>
</table>
</form>
</body>
</html>
<%
conn.close
set conn=nothing
%>