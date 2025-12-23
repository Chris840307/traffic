<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://10.104.10.246/traffic/smsx.cab#Version=6,1,432,1">
</object>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style3 {font-family:新細明體; color=0044ff; line-height:19px; font-size: 15px}
.style4 {font-family:新細明體;  line-height:19px;font-size: 12pt}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>送達清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
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

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing
%>
<%
	strwhere=request("SQLstr")


%>

</head>
<body>
<form name=myForm method="post">
<%	
	strUnitName2="select UnitName from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"'"
	set rsUnitName2=conn.execute(strUnitName2)
	if not rsUnitName2.eof then
		TitleUnitName2=trim(rsUnitName2("UnitName"))
	end if
	rsUnitName2.close
	set rsUnitName2=nothing

	strTitle="select Value from Apconfigure where ID=40"
	set rsTitle=conn.execute(strTitle)
	if not rsTitle.eof then
		TitleValue=rsTitle("Value")&" "&TitleUnitName2
		TitleValue=Replace(TitleValue,"台","臺")
	end if
	rsTitle.close
	set rsTitle=nothing
%>
<%		strwhere=""
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
			dciStr=" and BatchNumber in ('"&Sys_BatchNumber&"')"
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
		UserMarkDate1=gOutDT(request("Sys_UserMarkDate1"))&" 0:0:0"
		UserMarkDate2=gOutDT(request("Sys_UserMarkDate2"))&" 23:59:59"
		strwhere=strwhere&" and g.UserMarkDate between TO_DATE('"&UserMarkDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&UserMarkDate2&"','YYYY/MM/DD/HH24/MI/SS')"

		strCnt="select count(*) as cnt" &_
		" from BillBase a,(select billsn from DciLog where exchangetypeid='A'"&dciStr&") b,(select CarNo,Owner from billbasedcireturn where exchangetypeid='A') e,StopBillMailHistory g" &_
		" where a.Sn=b.BillSn and a.CarNo=e.CarNo and a.Sn=g.BillSn" &_
		" and a.RecordStateID=0 and g.UserMarkResonID in ('A','B','C')" &strwhere &_
		" order by g.UserMarkDate"

		set rsCnt=conn.execute(strCnt)
		if not rsCnt.eof then
			GetCnt=cdbl(rsCnt("Cnt"))
		end if
		rsCnt.close
		set rsCnt=nothing

		strSQL="select a.SN,a.BillNO,a.CarNO,e.Owner,a.CarSimpleID,a.IllegalDate" &_
		",a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.BillUnitID,a.BillMem1,a.BillMem2,a.ImageFileNameB,g.UserMarkResonID" &_
		" from BillBase a,(select billsn from DciLog where exchangetypeid='A'"&dciStr&") b,(select CarNo,Owner from billbasedcireturn where exchangetypeid='A') e,StopBillMailHistory g" &_
		" where a.Sn=b.BillSn and a.CarNo=e.CarNo and a.Sn=g.BillSn" &_
		" and a.RecordStateID=0 and g.UserMarkResonID in ('A','B','C')" &strwhere &_
		" order by g.UserMarkDate"

		set rs1=conn.execute(strSQL)
		If Not rs1.Bof Then rs1.MoveFirst 
		While Not rs1.Eof
		if PrintSN>0 then response.write "<div class=""PageNext""></div>"
%>		
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
	<td align="center" height="28" colspan="2"><span class="style4"><%
		
		response.write TitleValue&"&nbsp(收受)資料"

		if GetCnt="0" then
			pagecnt=1
		else
			pagecnt=fix(GetCnt/25+0.9999999)
		end if
	%></span></td>
	</tr>
	<tr>
	<td width="65%">
	列印日期：<%=now%>
	</td>
	<td align="right" width="35%">
	Page <%=fix(PrintSN/25)+1%> of <%=pagecnt%></td></td>
	</tr>
	</table>
	<table width="100%" border="<%
	if sys_City="嘉義縣" then
		response.write "1"
	else
		response.write "0"
	end if
	%>" cellpadding="1" cellspacing="0">
		<tr>
			<td>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td width="14%">單號</td>
					<td width="8%">違規日期</td>
					<td width="8%"></td>
					<td width="8%"></td>
					<td width="21%"></td>
					<td width="14%"></td>
					<td width="10%"></td>
					<td width="17%"></td>
				</tr>
				<tr>
					<td></td>
					<td>違規時間</td>
					<td>車號</td>
					<td>法條</td>
					<td>駕駛人/車主</td>
					<td>舉發單位</td>
					<td>員警</td>
					<td>送達原因</td>
				</tr>
			</table>
			</td>
		<tr>
<%		for i=1 to 25
			if rs1.eof then exit for
			PrintSN=PrintSN+1
%>
		<tr>
			<td>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td width="14%"><%
				'單號
				if trim(rs1("ImageFileNameB"))<>"" and not isnull(rs1("ImageFileNameB")) then
					response.write trim(rs1("ImageFileNameB"))
				else
					response.write "&nbsp;"
				end if
				%></td>
					<td width="8%"><%
					'違規日期
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gInitDT(rs1("IllegalDate"))
			else
				response.write "&nbsp;"
			end if
					%></td>
					<td width="8%"><%
				'車號
				if trim(rs1("CarNO"))<>"" and not isnull(rs1("CarNO")) then
					response.write trim(rs1("CarNO"))
				else
					response.write "&nbsp;"
				end if	
				%></td>
					<td width="8%"><%
				'法條一
				if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
					response.write trim(rs1("Rule1"))
				else
					response.write "&nbsp;"
				end if	
				%></td>
					<td width="21%"></td>
					<td width="14%"><%
					'舉發單位
			strUnit="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
			set rsUnit=conn.execute(strUnit)
			if not rsUnit.eof then
				response.write rsUnit("UnitName")
			else
				response.write "&nbsp;"
			end if
			rsUnit.close
			set rsUnit=nothing
					%></td>
					<td width="10%"><%
					'員警1
			if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
				response.write trim(rs1("BillMem1"))
			else
				response.write "&nbsp;"
			end if		
					%></td>
					<td width="17%" rowspan="2" valign="top"><%
					'退件原因
				strCode="select Content from DCIcode where TypeID=7 and ID='"&trim(rs1("UserMarkResonID"))&"'"
				set rsCode=conn.execute(strCode)
				if not rsCode.eof then
					response.write trim(rs1("UserMarkResonID"))&" "&trim(rsCode("Content"))
				end if
				rsCode.close
				set rsCode=nothing
				%></td>
				</tr>
				<tr>
					<td></td>
					<td><%
					'違規時間
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write Right("00"&hour(trim(rs1("IllegalDate"))),2)&Right("00"&minute(trim(rs1("IllegalDate"))),2)
			else
				response.write "&nbsp;"
			end if
					%></td>
					<td><%
					'車種
			if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
				response.write trim(rs1("CarSimpleID"))
			else
				response.write "&nbsp;"
			end if	
					%></td>
					<td><%
				'法條二
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
				'車主姓名
				if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
					response.write trim(rs1("Owner"))
				else
					response.write "&nbsp;"
				end if				
				%></td>
					<td></td>
					<td><%
					'員警2
			if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
				response.write trim(rs1("BillMem2"))
			else
				response.write "&nbsp;"
			end if		
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

</form>
</body>
</html>
<script language="javascript">
function DP(){
	window.focus();
	window.print();
}

printWindow(true,7,5.08,5.08,5.08);
</script>
<%conn.close%>