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
fname=year(now)&fMnoth&fDay&"_年度保留清冊清冊.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 

	strRul="select Value from Apconfigure where ID=3"
	set rsRul=conn.execute(strRul)
	RuleVer=trim(rsRul("Value"))
	rsRul.Close

	If Not ifnull(request("Sys_SendBillSN")) Then

		sys_billsn=request("Sys_SendBillSN")
	elseif Not ifnull(request("hd_BillSN")) Then

		sys_billsn=request("hd_BillSN")
	else

		sys_billsn=request("BillSN")
	End If 

	tmp_billsn=split(sys_billsn,",")

	sys_billsn=""

	For i = 0 to Ubound(tmp_billsn)

		If i >0 then

			If i mod 100 = 0 Then

				sys_billsn=sys_billsn&"@"
			elseif sys_billsn<>"" then

				sys_billsn=sys_billsn&","
			end If 
		end if

		sys_billsn=sys_billsn&tmp_billsn(i)

	Next

	tmpSQL=""

	If Ubound(tmp_billsn) >= 100 Then

		sys_billsn=split(sys_billsn,"@")
		
		For i = 0 to Ubound(sys_billsn)
			
			If tmpSQL <>"" Then tmpSQL=tmpSQL&" union all "
			
			tmpSQL=tmpSQL&"select sn from passerbase where sn in("&sys_billsn(i)&")"
		Next

	else

		tmpSQL="select sn from passerbase where sn in("&sys_billsn&")"

	End if 

	BasSQL="("&tmpSQL&") tmpPasser"
'檢查是否可進入本系統

	strSQL="select distinct a.SN,a.IllegalDate,a.BillNo,a.DriverID,a.Driver,a.BillFillDate,a.Rule1,a.Rule2,a.DeallineDate,(nvl(a.Forfeit1,0)+nvl(a.FORFEIT2,0)+nvl(a.FORFEIT3,0)+nvl(a.FORFEIT4,0)) Forfeit,b.JUDEDATE,b.OpenGovNumBer,f.PayDate,f.PayAmount,g.UnitName BillUnitName from PasserBase a,PasserJude b,PasserSend c,PasserUrge d,(select PasserSN,Max(ArrivedDate) ArrivedDate from PassersEndArrived group by PasserSN) e,(select BillSN,MAX(PayDate) PayDate,sum(PayAmount) PayAmount from PasserPay group by BillSN) f,(select UnitID,UnitName from UnitInfo) g where a.RecordStateID=0 and a.SN=b.BillSN(+) and a.BillNo=b.BillNo(+) and a.SN=c.BillSN(+) and a.BillNo=c.BillNo(+) and a.SN=d.BillSN(+) and a.BillNo=d.BillNo(+) and a.SN=e.PasserSN(+) and a.SN=f.BillSN(+) and a.BillUnitID=g.UnitID and f.PayDate is null and a.billstatus<>9 and Exists(select 'Y' from "&BasSQL&" where SN=a.SN) "

	strSQL=strSQL & " Union all "

	strSQL=strSQL & "select distinct a.SN,a.IllegalDate,a.BillNo,a.DriverID,a.Driver,a.BillFillDate,a.Rule1,a.Rule2,a.DeallineDate,(nvl(a.Forfeit1,0)+nvl(a.FORFEIT2,0)+nvl(a.FORFEIT3,0)+nvl(a.FORFEIT4,0)) Forfeit,b.JUDEDATE,b.OpenGovNumBer,f.PayDate,f.PayAmount,g.UnitName BillUnitName from PasserBase a,PasserJude b,PasserSend c,PasserUrge d,(select PasserSN,Max(ArrivedDate) ArrivedDate from PassersEndArrived group by PasserSN) e,(select BillSN,MAX(PayDate) PayDate,sum(PayAmount) PayAmount from PasserPay group by BillSN) f,(select UnitID,UnitName from UnitInfo) g where a.RecordStateID=0 and a.SN=b.BillSN(+) and a.BillNo=b.BillNo(+) and a.SN=c.BillSN(+) and a.BillNo=c.BillNo(+) and a.SN=d.BillSN(+) and a.BillNo=d.BillNo(+) and a.SN=e.PasserSN(+) and a.SN=f.BillSN(+) and a.BillUnitID=g.UnitID and b.JudeDate is not null and f.PayDate is not null and to_char(f.PayDate,'YYYY')> to_char(b.JudeDate,'YYYY') and Exists(select 'Y' from "&BasSQL&" where SN=a.SN) "

	tmpstrSQL="select * from ("&strSQL&") order by DriverID,IllegalDate"
	set rsfound=conn.execute(tmpstrSQL)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>年度保留清冊</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table width="700" border="1">
	<tr><td>
		<table width="100%" border="0" cellpadding="1" cellspacing="1">
			<tr>
				<td colspan="12" align="center"><strong>民國<%=year(rsfound("IllegalDate"))-1911%>年度保留清冊</strong></td>
			</tr>
			<tr>
				<td colspan="4" align="left">列印日期：<%=gInitDt(date)%></td>
			</tr>
			<tr>
				<td colspan="4" align="left">登入者：<%=Session("Ch_Name")%></td>
			</tr>
		</table>
	</td></tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="1" cellspacing="1">
<!--				<tr><td>
					<table width="100%" border="0" cellpadding="4" cellspacing="1"> -->
						<tr>
							<td>編號</td>
							<td>舉發單位</td>
							<td>單號</td>
							<td>違規人</td>
							<td>違規人ID</td>
							<td>舉發日期</td>
							<td>應到案日</td>
							<td>應納金額</td>
							<td>違規法條</td>
							<td>裁決日</td>
							<td>裁決文號</td>
							<td>備註</td>
						</tr>
<!--					</table>
				</td></tr>-->
				<%
				filecnt=0:sumFile=0:tmpDriverID="":sumNT=0:cntNt=0:Sys_Payamount=0

				If Not rsfound.eof Then	tmpDriverID=trim(rsfound("DriverID"))

				'response.write "<tr><td>"
				'response.write "<table width=""100%"" border=""0"" cellpadding=""4"" cellspacing=""1"">"
				while Not rsfound.eof
					sumFile=sumFile+1
					Sys_Payamount=0

					strSQL="select Level1,Level2 from law where version="&RuleVer&" and itemid='"&trim(rsfound("Rule1"))&"'"
					set rslaw=conn.execute(strSQL)

					response.write "<tr>"
					response.write "<td align=""left"">"&sumFile&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("BillUnitName"))&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("BillNo"))&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("Driver"))&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("DriverID"))&"</td>"
					response.write "<td align=""left"">"&trim(gInitDT(rsfound("BillFillDate")))&"</td>"
					response.write "<td align=""left"">"&trim(gInitDT(rsfound("DeallineDate")))&"</td>"
					response.write "<td>"&trim(rslaw("Level2"))&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("Rule1"))&"</td>"
					response.write "<td align=""left"">"&trim(gInitDT(rsfound("JudeDate")))&"</td>"
					response.write "<td>"&trim(rsfound("OpenGovNumBer"))&"</td>"


					response.write "<td align=""left""></td>"
					
					response.write "</tr>"

					sumNT=sumNT+cdbl(rslaw("Level1"))
					cntNt=cntNt+cdbl(rslaw("Level2"))
					rslaw.close
'					filecnt=filecnt+1
					rsfound.MoveNext
'					if Not rsfound.eof then
'						if trim(rsfound("DriverID"))<>trim(tmpDriverID) then
'							response.write "</table>"
'							response.write "</td></tr>"
'							response.write "<tr><td><table border=0><td align=""right"" colspan=""7"">&nbsp;</td>"
'							response.write "<td align=""right"">小計：</td>"
'							response.write "<td align=""right"">"&filecnt&"筆"&"</td>"
'							response.write "<td align=""right"">"&sumNT&"</td></tr></table></td></tr>"
'							response.write "<tr><td>"
'							response.write "<table width=""100%"" border=""0"" cellpadding=""4"" cellspacing=""1"">"
'							cntNt=cntNt+sumNT
'							tmpDriverID=trim(rsfound("DriverID"))
'							sumNT=0
'							filecnt=0
'						end if
'					end if
				wend
'				response.write "</table>"
'				response.write "</td></tr>"
'				response.write "<tr><td><table border=0><td align=""right"" colspan=""7"">&nbsp;</td>"
'				response.write "<td align=""right"">小計：</td>"
'				response.write "<td align=""right"">"&filecnt&"筆"&"</td>"
'				response.write "<td align=""right"">"&sumNT&"</td></tr></table></td></tr>"
'				cntNt=cntNt+sumNT
				'response.write "<tr><td><table border=0><td align=""right"" colspan=""7"">&nbsp;</td>"
				response.write "<tr><td colspan=""11""><table border=0><tr><td align=""right"">&nbsp;</td>"
				response.write "<td align=""right"">合計：</td>"
				'response.write "<td align=""right"">"&sumFile&"筆"&"</td>"
				response.write "<td align=""right"">應納金額："&cntNt&"</td></tr></table></td></tr>"
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