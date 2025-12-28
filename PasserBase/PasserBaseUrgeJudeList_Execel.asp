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
fname=year(now)&fMnoth&fDay&"_強制執行移送清冊.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 

'檢查是否可進入本系統
	strwhere=""
	if request("FromILLEGALDATE")<>"" and request("TOILLEGALDATE")<>""then
		ArgueDate1=gOutDT(request("FromILLEGALDATE"))&" 0:0:0"
		ArgueDate2=gOutDT(request("TOILLEGALDATE"))&" 23:59:59"
		strwhere=" and a.IllegalDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if
	strSQLTemp="select a.SN,a.IllegalDate,a.BillNo,a.Driver,a.DriverID,a.Rule1,a.FORFEIT1,a.BillFillDate,a.DeallineDate,b.SENDDATE,c.UnitName as BillUnitName,d.Payamount,e.JUDEDATE,e.OpenGovNumBer as JudeNo from PasserBase a,PasserSend b,UnitInfo c,PasserPay d,PasserJude e where a.RecordStateID=0 and a.SN=b.BillSN(+) and a.BillNo=b.BillNo(+) and a.BillUnitID=c.UnitID(+) and a.SN=d.BillSN(+) and a.BillNo=d.BillNo(+) and a.SN=e.BillSN(+) and a.BillNo=e.BillNo(+)"&strwhere&" and a.SN in("&trim(request("hd_BillSN"))&") order by DriverID"

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>強制執行移送清冊</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table width="700" border="1">
	<tr><td>
		<table width="100%" border="0" cellpadding="4" cellspacing="1">
			<tr>
				<td align="center"><strong>強制執行移送清冊</strong></td>
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
							<td>裁決日期/裁決案號</td>
							<td>金額</td>
						</tr>
					</table>
				</td></tr>
				<%
				set rsfound=conn.execute(strSQLTemp)
				filecnt=0:sumFile=0:tmpDriverID="":showDriver="":showDriverID="":sumNT=0:cntNt=0:Sys_Payamount=0
				tmpDriverID=trim(rsfound("DriverID"))
				showDriverID=trim(rsfound("DriverID"))
				showDriver=trim(rsfound("Driver"))))

				response.write "<tr><td>"
				response.write "<table width=""100%"" border=""0"" cellpadding=""4"" cellspacing=""1"">"
				while Not rsfound.eof
					sumFile=sumFile+1
					tmpDriverID=trim(rsfound("DriverID"))
					Sys_Payamount=trim(rsfound("FORFEIT1"))

					response.write "<tr><td align=""left"">"
					response.write trim(showDriverID)
					response.write "</td>"
					response.write "<td align=""left"">"&trim(showDriver)&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("BillNo"))&"</td>"
					response.write "<td align=""left"">"&trim(gInitDT(DateValue(rsfound("IllegalDate"))))&"</td>"
					response.write "<td align=""left"">"&trim(gInitDT(rsfound("BillFillDate")))&"</td>"
					response.write "<td align=""left"">"&trim(gInitDT(rsfound("DeallineDate")))&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("Rule1"))&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("BillUnitName"))&"</td>"
					response.write "<td align=""left"">"&trim(gInitDT(rsfound("JUDEDATE")))
					if trim(gInitDT(rsfound("JUDEDATE")))<>"" and trim(rsfound("JudeNo"))<>"" then
						response.write "／"
					end if
					response.write trim(rsfound("JudeNo"))&"</td>"
					response.write "<td>"&Sys_Payamount&"</td>"
					response.write "</tr>"

					if Not isnull(Sys_Payamount) then sumNT=sumNT+Cint(Sys_Payamount)
					filecnt=filecnt+1
					rsfound.MoveNext
					showDriver="":showDriverID=""
					if Not rsfound.eof then
						if trim(rsfound("DriverID"))<>trim(tmpDriverID) then
							showDriverID=trim(rsfound("DriverID"))
							showDriver=trim(rsfound("Driver"))
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