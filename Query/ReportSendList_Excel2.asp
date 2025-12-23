<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
Server.ScriptTimeout = 800
Response.flush
%>
<%
'權限
'AuthorityCheck(234)

 'and a.BillTypeID<>'2'
%>
<%
	StationArrayTemp=""
	strwhere=request("SQLstr")
	strStation="select distinct(e.DCIReturnStation) from DCILog a,MemberData b,DCIReturnStatus d" &_
		", BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and a.ExchangeTypeID='W' and a.DCIReturnStatusID='Y'" &_
		" and a.BillTypeID='2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0  "&strwhere
	set rsStation=conn.execute(strStation)
	If Not rsStation.Bof Then rsStation.MoveFirst 
	While Not rsStation.Eof
		if StationArrayTemp="" then
			StationArrayTemp=trim(rsStation("DCIReturnStation"))
		else
			StationArrayTemp=StationArrayTemp&","&trim(rsStation("DCIReturnStation"))
		end if
	rsStation.MoveNext
	Wend
	rsStation.close
	set rsStation=nothing
	strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d," &_
		" BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
		" and a.CarNo=e.CarNo and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DCIReturnStatusID=e.Status and a.ExchangeTypeID='W' and a.DCIReturnStatusID='Y'" &_
		" and a.BillTypeID='2'" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.RecordMemberID=b.MemberID(+) and f.RecordStateID=0  "&strwhere
	
	set rsCnt=conn.execute(strCnt)
	if not rsCnt.eof then
		DBcnt=rsCnt("Cnt")
	end if
	rsCnt.close
	set rsCnt=nothing
'response.write strSQL
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>逕行舉發移送清冊</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name=myForm method="post">
<%
	StationArray=split(StationArrayTemp,",")
	for SA=0 to ubound(StationArray)
%>
	<center><font size="3">逕行舉發移送清冊</font></center>
	舉發單位：交通隊&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	製表日期：<%=gInitDT(now)%>
	<br>
	應到案處所：<%
	strSqlStationName="select DCIstationName from Station where DCIstationID='"&trim(StationArray(SA))&"'"
	set rsSN=conn.execute(strSqlStationName)
	if not rsSN.eof then
		response.write trim(rsSN("DCIstationName"))
	end if
	rsSN.close
	set rsSN=nothing
	%>
	<table width="100%" border="1" cellpadding="1" cellspacing="0">
		<tr>
			<td width="2%" height="28" align="center">NO</td>
			<td width="7%" align="center">舉發單號</td>
			<td width="5%" align="center">違規日期</td>
			<td width="5%" align="center">違規時間</td>
			<td width="12%" align="center">違規地點</td>
			<td width="7%" align="center">詳細車種</td>
			<td width="7%" align="center">車牌號碼</td>
			<td width="7%" align="center">違反法條</td>
			<td width="7%" align="center">駕駛人 / 車主</td>
			<td width="8%" align="center">舉發單位</td>
			<td width="5%" align="center">員警</td>
			<td width="5%" align="center">扣件</td>
			<td width="5%" align="center">入案日期</td>
			<td width="5%" align="center">入案結果</td>
			<td width="5%" align="center">車籍資料</td>
			<td width="5%" align="center">駕籍資料</td>
		</tr>
<%		PrintSN=0
		strSQL="select f.SN,f.BillNo,f.BillTypeID,f.CarNo,f.IllegalDate,f.RecordDate" &_
			",e.DCIReturnCarType,f.Rule1,f.Rule2,f.Rule3,f.Rule4,f.Driver,e.DriverHomeZip" &_
			",e.DriverHomeAddress,f.DriverID,f.BillMem1,e.DCICaseInDate,e.DCIErrorCarData" &_
			",e.DCIErrorIDData,f.TrafficAccidentType,f.IllegalAddress,d.DCIReturnStatus" &_
			",e.Owner,a.BillUnitID from DCILog a,MemberData b,DCIReturnStatus d," &_
			"BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
			" and e.DCIReturnStation='"&trim(StationArray(SA))&"' and a.CarNo=e.CarNo" &_
			" and a.ExchangeTypeID=e.ExchangeTypeID and a.DCIReturnStatusID=e.Status" &_
			" and a.ExchangeTypeID='W' and a.DCIReturnStatusID='Y'" &_
			" and a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+)" &_
			" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+)" &_
			" and f.RecordStateID=0 "&strwhere
		set rs1=conn.execute(strSQL)
		If Not rs1.Bof Then rs1.MoveFirst 
		While Not rs1.Eof
%>		<tr>
			<td nowrap><%
			'No
			PrintSN=PrintSN+1
			response.write PrintSN
			%></td>
			<td nowrap><%
			'單號
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write trim(rs1("BillNO"))
			else
				response.write "&nbsp;"
			end if		
			%>
			</td>
			<td nowrap><%
			'違規日期
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gInitDT(rs1("IllegalDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td nowrap><%
			'違規時間
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write hour(rs1("IllegalDate"))&"："&minute(rs1("IllegalDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td nowrap><%
			'違規地點
			if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
				response.write trim(rs1("IllegalAddress"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td nowrap><%
			'詳細車種
			if trim(rs1("DCIReturnCarType"))<>"" and not isnull(rs1("DCIReturnCarType")) then
				strCType="select * from DCIcode where TypeID=5 and ID='"&trim(rs1("DCIReturnCarType"))&"'"
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
			<td nowrap><%
			'車牌號碼
			if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
				response.write trim(rs1("CarNo"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			'違反法條
			RuleStr=""
				if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
					RuleStr=trim(rs1("Rule1"))
				end if
				if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
					if RuleStr="" then
						RuleStr=trim(rs1("Rule2"))
					else
						RuleStr=RuleStr&" / "&trim(rs1("Rule2"))
					end if
				end if
				if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
					if RuleStr="" then
						RuleStr=trim(rs1("Rule3"))
					else
						RuleStr=RuleStr&" / "&trim(rs1("Rule3"))
					end if
				end if
				if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
					if RuleStr="" then
						RuleStr=trim(rs1("Rule4"))
					else
						RuleStr=RuleStr&" / "&trim(rs1("Rule4"))
					end if
				end if
				response.write RuleStr
			%></td>
			<td><%
			'駕駛人/車主
			'逕舉=2車籍資料 ,攔停駕籍資料
			if trim(rs1("BillTypeID"))="2" then
				response.write trim(rs1("Owner"))
			else
				response.write trim(rs1("Driver"))
			end if
			%></td>
			<td>
			<%
			'舉發單位
			if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
				strUName="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
				set rsUN=conn.execute(strUName)
				if not rsUN.eof then
					response.write trim(rsUN("UnitName"))
				end if
				rsUN.close
				set rsUN=nothing
			end if
			%></td>
			<td nowrap><%
			'員警
			if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
				response.write rs1("BillMem1")
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			'扣件
			strBillFastenerDetail="select Content from BillFastenerDetail a,DCIcode b where a.BillSN="&trim(rs1("SN"))&" and a.FastenerTypeID=b.ID and b.TypeID=6"
			set rsBF=conn.execute(strBillFastenerDetail)
			If Not rsBF.Bof Then
				rsBF.MoveFirst 
			else
				response.write "&nbsp;"
			end if
			While Not rsBF.Eof
				response.write rsBF("Content")
			rsBF.MoveNext
			Wend
			rsBF.close
			set rsBF=nothing
			%></td>
			<td nowrap><%
			'入案日期
			if trim(rs1("DCICaseInDate"))<>"" and not isnull(rs1("DCICaseInDate")) then
				response.write trim(rs1("DCICaseInDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td nowrap><%
			'入案結果
			if trim(rs1("DCIReturnStatus"))<>"" and not isnull(rs1("DCIReturnStatus")) then
				if trim(rs1("DCIReturnStatus"))="1" then
					response.write "成功"
				else
					response.write "異常"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td nowrap><%
			'車籍狀況
			if trim(rs1("DCIErrorCarData"))<>"" and not isnull(rs1("DCIErrorCarData")) then
				strCarData="select StatusContent from DCIReturnStatus where DCIActionID='WE' and DCIReturn='"&trim(rs1("DCIErrorCarData"))&"'"
				set rsCD=conn.execute(strCarData)
				if not rsCD.eof then
					response.write trim(rsCD("StatusContent"))
				else
					response.write "&nbsp;"
				end if
				rsCD.close
				set rsCD=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td nowrap><%
			'駕籍
			if trim(rs1("DCIErrorIDData"))<>"" and not isnull(rs1("DCIErrorIDData")) then
				strDriverData="select StatusContent from DCIReturnStatus where DCIActionID='WE' and DCIReturn='"&trim(rs1("DCIErrorIDData"))&"'"
				set rsDD=conn.execute(strDriverData)
				if not rsDD.eof then
					response.write trim(rsDD("StatusContent"))
				else
					response.write "&nbsp;"
				end if
				rsDD.close
				set rsDD=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>

		</tr>
<%			
		rs1.MoveNext
		Wend
		rs1.close
		set rs1=nothing

%>
	</table>
	共計： <%=PrintSN%>  &nbsp;筆<br>
<%if SA<>ubound(StationArray) then%>
	<div class="PageNext"></div>
<%end if
	next
%>
	<center>
	<input type="button" name="Print1" onclick="DP();" value="列印">
	</center>
</form>
</body>
</html>
<script language="javascript">
function DP(){
	window.focus();
	window.print();
}
</script>
<%conn.close%>