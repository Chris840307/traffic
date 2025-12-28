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

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
	

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

	strSQL="select a.SN,a.IllegalDate,a.BillNo,a.DriverID,a.Driver," &_
			"a.BillFillDate,a.Rule1,a.Rule2,a.DeallineDate,a.BillStatus," &_
			"(nvl(a.Forfeit1,0)+nvl(a.FORFEIT2,0)+nvl(a.FORFEIT3,0)+nvl(a.FORFEIT4,0)) Forfeit," &_
			"(Select JudeDate from PasserJude where billsn=a.sn) JudeDate," &_
			"(select MAX(PayDate) PayDate from PasserPay where billsn=a.sn) PayDate," &_
			"(select UnitName PayDate from UnitInfo where UnitID=a.BillUnitID) BillUnitName" &_
			" from PasserBase a where Exists(select 'Y' from "&BasSQL&" where SN=a.SN)"

	If Request("Sys_ReserveYear1") = "" Then 
		strSQL=strSQL&" and exists(select 'Y' from PasserJude where billsn=a.sn) and not exists(select 'Y' from PasserPay where to_char(PayDate,'YYYY')>(select to_char(JudeDate,'YYYY') from PasserJude where BillSN=PasserPay.BillSN) and billsn=a.sn) and a.billstatus<>9"

		strSQL=strSQL & " Union all "

		strSQL=strSQL & "select a.SN,a.IllegalDate,a.BillNo,a.DriverID,a.Driver," &_
				"a.BillFillDate,a.Rule1,a.Rule2,a.DeallineDate,a.BillStatus," &_
				"(nvl(a.Forfeit1,0)+nvl(a.FORFEIT2,0)+nvl(a.FORFEIT3,0)+nvl(a.FORFEIT4,0)) Forfeit," &_
				"(Select JudeDate from PasserJude where billsn=a.sn) JudeDate," &_
				"(select MAX(PayDate) PayDate from PasserPay where billsn=a.sn) PayDate," &_
				"(select UnitName PayDate from UnitInfo where UnitID=a.BillUnitID) BillUnitName" &_
				" from PasserBase a where Exists(select 'Y' from "&BasSQL&" where SN=a.SN)"

		If Request("Sys_ReserveYear1") = "" Then strSQL=strSQL&" and exists(select 'Y' from PasserJude where billsn=a.sn) and exists(select 'Y' from PasserPay where to_char(PayDate,'YYYY')>(select to_char(JudeDate,'YYYY') from PasserJude where BillSN=PasserPay.BillSN) and billsn=a.sn)"

	end if

	tmpstrSQL="select * from ("&strSQL&") order by DriverID,IllegalDate"

	set rsfound=conn.execute(tmpstrSQL)

	chkdatacnt=1
	If (sys_City = "彰化縣" and cdbl(Session("User_ID"))=549) or (sys_City = "基隆市" and Month(date)=1 ) Then
		
		If Request("Sys_ReserveYear1") ="" Then
			
			cntSQL="select count(1) cnt from PasserBase a where Exists(select 'Y' from "&BasSQL&" where SN=a.SN) and exists(select 'Y' from PasserJude where billsn=a.sn) and a.billstatus<>9 and a.ReserveYear is not null"
			set rscnt=conn.execute(cntSQL)

			If cdbl(rscnt("cnt"))=0 Then 
				chkdatacnt=0
				sys_year2=""

				strSQL2="select (Select to_char(min(JUDEDATE),'YYYY') from PasserJude where billsn=a.sn) JudeDate from PasserBase a where Exists(select 'Y' from "&BasSQL&" where SN=a.SN) and exists(select 'Y' from PasserJude where billsn=a.sn) and a.billstatus<>9 "&strwhere

				set rscnt2=conn.execute(strSQL2)
				
				sys_year2=cdbl(rscnt2("JUDEDATE"))-1911

				rscnt2.close
			end if

			rscnt.close
		
		end If 
	
	End if 

	If rsfound.eof Then
		Response.Write "查無資料。"
		Response.End
	End if 
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
		<table width="100%" border="0" cellpadding="4" cellspacing="1">
			<tr>
				<td align="center"><strong>民國<%=year(rsfound("JUDEDATE"))-1911%>年度保留清冊</strong></td>
			</tr>
			<tr>
				<td align="left">列印日期：<%=gInitDt(date)%></td>
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
							<td>單號</td>
							<td>違規人</td>
							<td>違規人ID</td>
							<td>違規日期</td>
							<td>填單日期</td>
							<td>應到案日</td>
							<td>法條一</td>
							<td>法條二</td>
							<td>舉發單位</td>
							<td>裁決日</td>
							<td>繳款日</td>
							<td>金額</td>
						</tr>
					</table>
				</td></tr>
				<%
				filecnt=0:sumFile=0:tmpDriverID="":sumNT=0:cntNt=0:Sys_Payamount=0

				If Not rsfound.eof Then	tmpDriverID=trim(rsfound("DriverID"))

				response.write "<tr><td>"
				response.write "<table width=""100%"" border=""0"" cellpadding=""4"" cellspacing=""1"">"
				while Not rsfound.eof
					sumFile=sumFile+1
					Sys_Payamount=0

					If (sys_City = "彰化縣" and cdbl(Session("User_ID"))=549) or (sys_City = "基隆市" and Month(date)=1 ) Then
					
						If chkdatacnt=0 Then

							strSQL="update PasserBase set ReserveYear="&sys_year2&" where sn="&rsfound("SN")&" and ReserveYear is null"

							conn.execute(strSQL)
						end If 
					end If 

					response.write "<tr>"
					response.write "<td align=""left"">"&trim(rsfound("BillNo"))&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("Driver"))&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("DriverID"))&"</td>"					
					response.write "<td align=""left"">"&trim(gInitDT(DateValue(rsfound("IllegalDate"))))&"</td>"
					response.write "<td align=""left"">"&trim(gInitDT(rsfound("BillFillDate")))&"</td>"
					response.write "<td align=""left"">"&trim(gInitDT(rsfound("DeallineDate")))&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("Rule1"))&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("Rule2"))&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("BillUnitName"))&"</td>"
					response.write "<td align=""left"">"&trim(gInitDT(rsfound("JudeDate")))&"</td>"
					response.write "<td align=""left"">"&trim(gInitDT(rsfound("PayDate")))&"</td>"

					Sys_PayAmount=0
					If trim(rsfound("BillStatus"))<>"9" Then						
						strSQL="select Nvl(sum(PayAmount),0) PayAmount from PasserPay where to_char(PayDate,'YYYY')>=(select to_char(JudeDate,'YYYY') from PasserJude where BillSN=PasserPay.BillSN) and BillSN="&trim(rsfound("SN"))
						set rspay=conn.execute(strSQL)
						If not rspay.eof Then
							Sys_PayAmount=cdbl(rspay("PayAmount"))
						End if
						rspay.close
					End if
					
					response.write "<td>"&cdbl(rsfound("Forfeit"))-Sys_PayAmount&"</td>"
					response.write "</tr>"

					if Not isnull(rsfound("Forfeit")) then sumNT=sumNT+cdbl(rsfound("Forfeit"))-Sys_PayAmount
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