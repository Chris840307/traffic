<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_催告催繳逾期已到案裁決清冊.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
	strwhere=""
	if trim(request("Sys_SQL"))<>"" then
		strwhere=request("Sys_SQL")
	else
		ArgueDate1=gOutDT(request("FromILLEGALDATE"))&" 0:0:0"
		ArgueDate2=gOutDT(request("TOILLEGALDATE"))&" 23:59:59"
		strwhere=" and a.IllegalDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if
'檢查是否可進入本系統
	
	BasSQL="select distinct a.SN from PasserBase a,PasserJude b,PasserSend c,PasserUrge d,PassersEndArrived e,(select distinct BillSN,PayDate from PasserPay) f where a.RecordStateID=0 and a.SN=b.BillSN(+) and a.BillNo=b.BillNo(+) and a.SN=c.BillSN(+) and a.BillNo=c.BillNo(+) and a.SN=d.BillSN(+) and a.BillNo=d.BillNo(+) and a.SN=e.PasserSN(+) and a.SN=f.BillSN(+) and a.RecorDStateID<>-1 "&strwhere

	strSQLTemp="select a.SN,a.IllegalDate,a.BillNo,a.Driver,a.DriverID,a.Rule1,a.BillFillDate,a.DeallineDate,b.SENDDATE,c.UnitName as BillUnitName from PasserBase a,PasserSend b,UnitInfo c,PasserUrge d where a.RecordStateID=0 and a.SN=b.BillSN(+) and a.BillNo=b.BillNo(+) and a.BillUnitID=c.UnitID(+) and a.SN=d.BillSN(+) and a.BillNo=d.BillNo(+) and Not(d.URGEDATE is null) and a.BILLSTATUS=9 and a.SN in("&BasSQL&") order by DriverID"

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>催告催繳逾期已到案裁決清冊</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table width="700" border="1">
	<tr><td>
		<table width="100%" border="0" cellpadding="4" cellspacing="1">
			<tr>
				<td align="center"><strong>催告催繳逾期已到案裁決清冊</strong></td>
			</tr>
			<tr>
				<td align="left">列印日期：<%=gInitDt(date)%></td>
			</tr>
			<tr>
				<td align="left">處理時間：<%=request("FromILLEGALDATE")%><%if trim(request("FromILLEGALDATE"))<>"" and trim(request("TOILLEGALDATE"))<>"" then response.write "～"%><%=request("TOILLEGALDATE")%></td>
			</tr>
			<tr>
				<td align="left">登入者：<%=Session("Ch_Name")%></td>
			</tr>
		</table>
	</td></tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="1">
				<tr><td>
					<table width="100%" border="0" cellpadding="4" cellspacing="1">
						<tr>
							<td>違規人ID</td>
							<td>違規人</td>
							<td>單號</td>
							<td>違規日期</td>
							<td>填單日期</td>
							<td>應到案日</td>
							<td>法條一</td>
							<td>舉發單位</td>
							<td>強制移送日期</td>
							<td>金額</td>
						</tr>
					</table>
				</td></tr>
				<%
				set rsfound=conn.execute(strSQLTemp)
				filecnt=0:sumFile=0:tmpDriverID="":sumNT=0:cntNt=0:Sys_Payamount=0
				
				if Not rsfound.eof then
					strSQL="select sum(Payamount) as Sys_Payamount from PasserPay where BillSN="&rsfound("SN")&" and BillNo='"&rsfound("BillNo")&"'"
					set rspay=conn.execute(strSQL)
				
					if not rspay.eof then Sys_Payamount=rspay("Sys_Payamount")
					rspay.close
				end if

				if Not rsfound.eof then
					tmpDriverID=trim(rsfound("DriverID"))
					'if Not isnull(Sys_Payamount) then sumNT=Cint(Sys_Payamount)
				end if
				response.write "<tr><td>"
				response.write "<table width=""100%"" border=""0"" cellpadding=""4"" cellspacing=""1"">"
				while Not rsfound.eof
					sumFile=sumFile+1
					Sys_Payamount=0
					strSQL="select sum(Payamount) as Sys_Payamount from PasserPay where BillSN="&rsfound("SN")&" and BillNo='"&rsfound("BillNo")&"'"
					set rspay=conn.execute(strSQL)
					if not rspay.eof then
						if Not isnull(rspay("Sys_Payamount")) then Sys_Payamount=rspay("Sys_Payamount")
					end if
					rspay.close

					response.write "<tr>"
					response.write "<td align=""left"">"&trim(rsfound("DriverID"))&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("Driver"))&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("BillNo"))&"</td>"
					response.write "<td align=""left"">"&trim(gInitDT(DateValue(rsfound("IllegalDate"))))&"</td>"
					response.write "<td align=""left"">"&trim(gInitDT(rsfound("BillFillDate")))&"</td>"
					response.write "<td align=""left"">"&trim(gInitDT(rsfound("DeallineDate")))&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("Rule1"))&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("BillUnitName"))&"</td>"
					response.write "<td align=""left"">"&trim(gInitDT(rsfound("SendDate")))&"</td>"
					response.write "<td>"&Sys_Payamount&"</td>"
					response.write "</tr>"

					if Not isnull(Sys_Payamount) then sumNT=sumNT+Cint(Sys_Payamount)
					filecnt=filecnt+1
					rsfound.MoveNext
					if Not rsfound.eof then
						if trim(rsfound("DriverID"))<>trim(tmpDriverID) then
							response.write "</table>"
							response.write "</td></tr>"
							response.write "<tr><td><table border=0><td align=""right"" colspan=""7"">&nbsp;</td>"
							response.write "<td align=""right"">小計：</td>"
							response.write "<td align=""right"">"&filecnt&"筆"&"</td>"
							response.write "<td align=""right"">"&sumNT&"</td></tr></table></td></tr>"
							response.write "<tr><td>"
							response.write "<table width=""100%"" border=""0"" cellpadding=""4"" cellspacing=""1"">"
							cntNt=cntNt+sumNT
							tmpDriverID=trim(rsfound("DriverID"))
							sumNT=0
							filecnt=0
						end if
					end if
				wend
				response.write "</table>"
				response.write "</td></tr>"
				response.write "<tr><td><table border=0><td align=""right"" colspan=""7"">&nbsp;</td>"
				response.write "<td align=""right"">小計：</td>"
				response.write "<td align=""right"">"&filecnt&"筆"&"</td>"
				response.write "<td align=""right"">"&sumNT&"</td></tr></table></td></tr>"
				cntNt=cntNt+sumNT
				response.write "<tr><td><table border=0><td align=""right"" colspan=""7"">&nbsp;</td>"
				response.write "<td align=""right"">共計：</td>"
				response.write "<td align=""right"">"&sumFile&"筆"&"</td>"
				response.write "<td align=""right"">"&cntNt&"</td></tr></table></td></tr>"
				rsfound.close
				set rsfound=nothing
				%>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
<%
conn.close
set conn=nothing
%>