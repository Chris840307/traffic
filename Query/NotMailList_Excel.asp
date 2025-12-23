<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing
%>
<%if sys_City<>"雲林縣" and sys_City<>"台中縣" and sys_City<>"嘉義縣" then%>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://10.104.10.246/traffic/smsx.cab#Version=6,1,432,1">
</object>
<%end if%>
<%
Server.ScriptTimeout = 800
Response.flush
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style3 {font-family:新細明體; color=0044ff; line-height:19px; font-size: 15px}
.style4 {font-family:新細明體; color=0044ff; line-height:12px; font-size: 10px}
.style5 {font-family:新細明體; color=0044ff; line-height:13px; font-size: 11px}
.style6 {font-family:新細明體; color=0044ff; line-height:12px; font-size: 10px}
<%if sys_City="雲林縣" or sys_City="台中縣" or sys_City="嘉義縣" then%>
.pageprint {
  margin-left: 7mm;
  margin-right: 5.08mm;
  margin-top: 5.08mm;
  margin-bottom: 5.08mm;
}
<%end if%>
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>不郵寄清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<%
'權限
'AuthorityCheck(234)

 'and a.BillTypeID<>'2'
%>
<%


	'頁數
	PageNum=1
	StationArrayTemp=""
	strwhere=request("SQLstr")

%>
</head>
<body>
<form name=myForm method="post">
<%
	PrintSNtotal=0	'編號
	if sys_City="嘉義縣" or sys_City="南投縣" or sys_City="台南市" or sys_City="宜蘭縣" then
		PageCount=20
	else
		PageCount=23
	end if

	'其他堅理所列表
		PrintSN=0
if sys_City="基隆市" then
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d," &_
			"BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
			" and a.CarNo=e.CarNo and EquipMentID=-1" &_
			" and a.ExchangeTypeID=e.ExchangeTypeID and a.DCIReturnStatusID=e.Status" &_
			" and a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('Y','S','n')" &_
			" and ((((((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','L','T')) or (e.DciErrorCarData='F' and e.rule4='2607')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
			" and a.BillTypeID='2')" &_
			" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n')))" &_
			" and a.ExchangeTypeID=d.DCIActionID(+)" &_
			" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+)" &_
			" and f.RecordStateID=0 "&strwhere
elseif sys_City="台南市" then
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d," &_
			"BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
			" and a.CarNo=e.CarNo and EquipMentID=-1" &_
			" and a.ExchangeTypeID=e.ExchangeTypeID and a.DCIReturnStatusID=e.Status" &_
			" and a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('Y','S','n')" &_
			" and (((((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
			" and a.BillTypeID='2')" &_
			" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n')))" &_
			" and a.ExchangeTypeID=d.DCIActionID(+)" &_
			" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+)" &_
			" and f.RecordStateID=0 "&strwhere
elseif sys_City="台南縣" or sys_City="雲林縣" then
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d," &_
			"BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
			" and a.CarNo=e.CarNo and EquipMentID=-1" &_
			" and a.ExchangeTypeID=e.ExchangeTypeID and a.DCIReturnStatusID=e.Status" &_
			" and a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('Y','S','n')" &_
			" and ((((d.DCIreturnStatus=1 and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
			" and a.BillTypeID='2')" &_
			" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n')))" &_
			" and a.ExchangeTypeID=d.DCIActionID(+)" &_
			" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+)" &_
			" and f.RecordStateID=0 "&strwhere
else
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d," &_
			"BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
			" and a.CarNo=e.CarNo and EquipMentID=-1" &_
			" and a.ExchangeTypeID=e.ExchangeTypeID and a.DCIReturnStatusID=e.Status" &_
			" and a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('Y','S','n')" &_
			" and (((((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
			" and a.BillTypeID='2')" &_
			" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n')))" &_
			" and a.ExchangeTypeID=d.DCIActionID(+)" &_
			" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+)" &_
			" and f.RecordStateID=0 "&strwhere
end if
		set rsCnt=conn.execute(strCnt)
		if not rsCnt.eof then
			if trim(rsCnt("cnt"))="0" then
				pagecnt=1
			else
				pagecnt=fix(Cint(rsCnt("cnt"))/PageCount+0.9999999)
			end if
		end if
		rsCnt.close
		set rsCnt=nothing
if sys_City="基隆市" then
		strSQL="select f.SN,f.BillNo,f.BillTypeID,f.CarNo,f.CarSimpleID,f.IllegalDate,f.RecordDate" &_
			",e.DCIReturnCarType,f.Rule1,f.Rule2,f.Rule3,f.Rule4,e.Driver,e.DriverHomeZip" &_
			",e.DriverHomeAddress,f.DriverID,f.BillMem1,e.DCICaseInDate,e.DCIErrorCarData" &_
			",e.DCIErrorIDData,f.TrafficAccidentType,f.IllegalAddress" &_
			",d.DCIReturnStatus,a.FileName,a.BatchNumber" &_
			",e.Owner,a.BillUnitID from DCILog a,MemberData b,DCIReturnStatus d," &_
			"BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
			" and a.CarNo=e.CarNo and EquipMentID=-1" &_
			" and a.ExchangeTypeID=e.ExchangeTypeID and a.DCIReturnStatusID=e.Status" &_
			" and a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('Y','S','n')" &_
			" and ((((((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','L','T')) or (e.DciErrorCarData='F' and e.rule4='2607')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
			" and a.BillTypeID='2')" &_
			" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n')))" &_
			" and a.ExchangeTypeID=d.DCIActionID(+)" &_
			" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+)" &_
			" and f.RecordStateID=0 "&strwhere&" order by f.RecordMemberID,f.RecordDate"
elseif sys_City="台南市" then
		strSQL="select f.SN,f.BillNo,f.BillTypeID,f.CarNo,f.CarSimpleID,f.IllegalDate,f.RecordDate" &_
			",e.DCIReturnCarType,f.Rule1,f.Rule2,f.Rule3,f.Rule4,e.Driver,e.DriverHomeZip" &_
			",e.DriverHomeAddress,f.DriverID,f.BillMem1,e.DCICaseInDate,e.DCIErrorCarData" &_
			",e.DCIErrorIDData,f.TrafficAccidentType,f.IllegalAddress" &_
			",d.DCIReturnStatus,a.FileName,a.BatchNumber" &_
			",e.Owner,a.BillUnitID from DCILog a,MemberData b,DCIReturnStatus d," &_
			"BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
			" and a.CarNo=e.CarNo and EquipMentID=-1" &_
			" and a.ExchangeTypeID=e.ExchangeTypeID and a.DCIReturnStatusID=e.Status" &_
			" and a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('Y','S','n')" &_
			" and (((((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
			" and a.BillTypeID='2')" &_
			" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n')))" &_
			" and a.ExchangeTypeID=d.DCIActionID(+)" &_
			" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+)" &_
			" and f.RecordStateID=0 "&strwhere&" order by f.RecordMemberID,f.RecordDate"
elseif sys_City="台南縣" or sys_City="雲林縣" then
		strSQL="select f.SN,f.BillNo,f.BillTypeID,f.CarNo,f.CarSimpleID,f.IllegalDate,f.RecordDate" &_
			",e.DCIReturnCarType,f.Rule1,f.Rule2,f.Rule3,f.Rule4,e.Driver,e.DriverHomeZip" &_
			",e.DriverHomeAddress,f.DriverID,f.BillMem1,e.DCICaseInDate,e.DCIErrorCarData" &_
			",e.DCIErrorIDData,f.TrafficAccidentType,f.IllegalAddress" &_
			",d.DCIReturnStatus,a.FileName,a.BatchNumber" &_
			",e.Owner,a.BillUnitID from DCILog a,MemberData b,DCIReturnStatus d," &_
			"BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
			" and a.CarNo=e.CarNo and EquipMentID=-1" &_
			" and a.ExchangeTypeID=e.ExchangeTypeID and a.DCIReturnStatusID=e.Status" &_
			" and a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('Y','S','n')" &_
			" and ((((d.DCIreturnStatus=1 and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
			" and a.BillTypeID='2')" &_
			" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n')))" &_
			" and a.ExchangeTypeID=d.DCIActionID(+)" &_
			" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+)" &_
			" and f.RecordStateID=0 "&strwhere&" order by f.RecordMemberID,f.RecordDate"
else
		strSQL="select f.SN,f.BillNo,f.BillTypeID,f.CarNo,f.CarSimpleID,f.IllegalDate,f.RecordDate" &_
			",e.DCIReturnCarType,f.Rule1,f.Rule2,f.Rule3,f.Rule4,e.Driver,e.DriverHomeZip" &_
			",e.DriverHomeAddress,f.DriverID,f.BillMem1,e.DCICaseInDate,e.DCIErrorCarData" &_
			",e.DCIErrorIDData,f.TrafficAccidentType,f.IllegalAddress" &_
			",d.DCIReturnStatus,a.FileName,a.BatchNumber" &_
			",e.Owner,a.BillUnitID from DCILog a,MemberData b,DCIReturnStatus d," &_
			"BillBaseDCIReturn e,BillBase f where f.SN=a.BillSN and a.BillNo=e.BillNO" &_
			" and a.CarNo=e.CarNo and EquipMentID=-1" &_
			" and a.ExchangeTypeID=e.ExchangeTypeID and a.DCIReturnStatusID=e.Status" &_
			" and a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('Y','S','n')" &_
			" and (((((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','T')) and UseTool<>8) or (d.DCIreturnStatus=1 and UseTool=8))" &_
			" and a.BillTypeID='2')" &_
			" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n')))" &_
			" and a.ExchangeTypeID=d.DCIActionID(+)" &_
			" and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+)" &_
			" and f.RecordStateID=0 "&strwhere&" order by f.RecordMemberID,f.RecordDate"
end if
		set rs1=conn.execute(strSQL)
		If Not rs1.Bof Then rs1.MoveFirst 
		While Not rs1.Eof

		if PrintSN>0 then
%>
		<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
<%
			response.write "<div class=""PageNext""></div>"
		end if
%>
	<table width="710" border="0" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center" colspan="2"><font size="3"><%=TitleUnitName%>&nbsp;不郵寄清冊</font></td>
		</tr>
		<tr>
		<td align="left">
		移送日期：<%=Right("00"&year(now)-1911,2)&Right("00"&month(now),2)&Right("00"&day(now),2)%>&nbsp; &nbsp; &nbsp;(本批案件已透過中華電信數據分公司作入案管制)
		</td>
		<td align="right">
		Page <%=fix(PrintSN/PageCount)+1%> of <%=pagecnt%>
		</td>
		</tr>
	</table>
	<table width="710" border="1" cellpadding="1" cellspacing="0">
	<tr>
	<td>
	<table width="710" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="5%"></td>
			<td width="10%">單號</td>
			<td width="9%">違規日期</td>
			<td width="9%"></td>
			<td width="8%"></td>
			<td width="18%"></td>
			<td width="11%">舉發單位</td>
			<td width="9%">員警</td>
			<td width="10%">扣件</td>
			<td width="11%">備註</td>
		</tr>
		<tr>
			<td>編號</td>
			<td>入案日期</td>
			<td>違規時間</td>
			<td>車號</td>
			<td>法條</td>
			<td>駕駛人/車主</td>
			<td>駕籍資料</td>
			<td></td>
			<td>車籍資料</td>
			<td></td>
		</tr>
	</table>
	</td>
	</tr>
<%		for i=1 to PageCount
			if rs1.eof then exit for
%>
	<tr>
	<td>
	<table width="710" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="5%"><%
			PrintSNtotal=PrintSNtotal+1
			PrintSN=PrintSN+1
			response.write PrintSNtotal
			%></td>
			<td width="10%"><%
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write rs1("BillNO")
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="9%"><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gInitDT(rs1("IllegalDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="9%"><%response.write trim(rs1("CarNo"))%></td>
			<td width="8%"><%
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="18%"><%
			if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
				response.write funcCheckFont(rs1("Driver"),18,1)
			else
				response.write "&nbsp;"
			end if	
			%></td>
			<td width="11%"><span class="style6"><%
			if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
				strUnit="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
				set rsUnit=conn.execute(strUnit)
				if not rsUnit.eof then
					response.write trim(rsUnit("UnitName"))
				end if
				rsUnit.close
				set rsUnit=nothing
			end if
			%></span></td>
			<td width="9%"><%
			if (trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1"))) then
				response.write rs1("BillMem1")
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td width="10%"><%
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
			<td width="11%"><%
			'檔名
			response.write "<span class='style4'>"&trim(rs1("FileName"))&"</span>"
			%></td>
		</tr>
		<tr>
			<td></td>
			<td><%
			if trim(rs1("DCICaseInDate"))<>"" and not isnull(rs1("DCICaseInDate")) then
				response.write trim(rs1("DCICaseInDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write Right("00"&hour(trim(rs1("IllegalDate"))),2)&Right("00"&minute(trim(rs1("IllegalDate"))),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%response.write trim(rs1("CarSimpleID"))%></td>
			<td><%
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				response.write trim(rs1("Rule2"))
			else
				response.write "&nbsp;"
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				response.write "<br>"&trim(rs1("Rule3"))
			end if
			%></td>
			<td><%
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
				response.write funcCheckFont(rs1("Owner"),18,1)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><span class="style5"><%
			'駕籍
			if trim(rs1("DCIErrorIDData"))="0" then
				response.write "0 正常"
			elseif trim(rs1("DCIErrorIDData"))<>"" and not isnull(rs1("DCIErrorIDData")) then
				strDriverData="select StatusContent from DCIReturnStatus where DCIActionID='WE' and DCIReturn='"&trim(rs1("DCIErrorIDData"))&"'"
				set rsDD=conn.execute(strDriverData)
				if not rsDD.eof then
					response.write trim(rs1("DCIErrorIDData"))&" "&trim(rsDD("StatusContent"))
				else
					response.write "&nbsp;"
				end if
				rsDD.close
				set rsDD=nothing
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td></td>
			<td><%
			'車籍狀況
			if trim(rs1("DCIErrorCarData"))="0" then
					response.write "0 正常"
			elseif trim(rs1("DCIErrorCarData"))<>"" and not isnull(rs1("DCIErrorCarData")) then
				strCarData="select StatusContent from DCIReturnStatus where DCIActionID='WE' and DCIReturn='"&trim(rs1("DCIErrorCarData"))&"'"
				set rsCD=conn.execute(strCarData)
				if not rsCD.eof then
					response.write trim(rs1("DCIErrorCarData"))&" "&trim(rsCD("StatusContent"))
				else
					response.write "&nbsp;"
				end if
				rsCD.close
				set rsCD=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td><%
			'批號
			response.write "<span class='style4'>"&trim(rs1("BatchNumber"))&"</span>"
			%></td>
		</tr>
		</table>
		</td>
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

%>
	共計： <%=PrintSN%>  &nbsp;筆<br>
		<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
</form>
</body>
</html>
<script language="javascript">
function DP(){
	window.focus();
	window.print();
}

<%if sys_City="雲林縣" or sys_City="台中縣" or sys_City="嘉義縣" then%>
window.print();
<%else%>
printWindow(true,7,5.08,5.08,5.08);
<%end if%>
</script>
<%conn.close%>