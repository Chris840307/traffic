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
fname=year(now)&fMnoth&fDay&"_慢車行人道路障礙違規舉發清冊.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 

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

	strSQLTemp="select a.SN,a.IllegalDate,a.BillNo,a.Driver,a.DriverID," &_
				"a.DriverBirth,a.DriverAddress,a.IllegalAddress,a.Rule1," &_
				"a.Forfeit1,a.BillMem1,a.BillFillDate,a.DeallineDate," &_
				"a.RecordDate,a.BILLSTATUS," &_
				"(Select JudeDate from PasserJude where billsn=a.sn) JUDEDATE," &_
				"(Select OpenGovNumBer from PasserJude where billsn=a.sn) JudeNo," &_
				"(Select SendDate from PasserSend where billsn=a.sn) SendDate," &_
				"(Select UrgeDate from PasserUrge where billsn=a.sn) UrgeDate," &_
				"(Select OpenGovNumBer from PasserUrge where billsn=a.sn) UrgeNo," &_
				"(Select UnitName from UnitInfo where UnitID=a.MemberStation) StationName," &_
				"(Select UnitName from UnitInfo where UnitID=a.BillUnitID) BillUnitName," &_
				"(Select sum(Payamount) from PasserPay where billsn=a.sn) Payamount," &_
				"(Select max(PayDate) from PasserPay where billsn=a.sn) PayDate," &_
				"(Select max(PayTypeID) from PasserPay where billsn=a.sn) PayTypeID," &_
				"(Select max(PayNo) from PasserPay where billsn=a.sn) PayNo" &_
				" from PasserBase a where a.RecordStateID=0 and Exists(select 'Y' from "&BasSQL&" where SN=a.SN)"

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>慢車/行人/道路障礙違規舉發清冊</title>
<style type="text/css">
<!--
.style1 {font-family: "新細明體"; font-size: 12px;}
-->
</style>
</head>
<body>
<table width="645" border="1">
	<tr><td>
		<table width="100%" border="0" cellpadding="4" cellspacing="1">
			<tr>
				<td class="style1" align="center"><strong>慢車/行人/道路障礙違規舉發清冊</strong></td>
			</tr>
			<tr>
				<td class="style1" align="left">列印日期：<%=gInitDt(date)%></td>
			</tr>
			<tr>
				<td class="style1" align="left">處理時間：<%=request("FromILLEGALDATE")%><%if trim(request("FromILLEGALDATE"))<>"" and trim(request("TOILLEGALDATE"))<>"" then response.write "～"%><%=request("TOILLEGALDATE")%></td>
			</tr>
			<tr>
				<td class="style1" align="left">登入者：<%=Session("Ch_Name")%></td>
			</tr>
		</table>
	</td></tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="1">
				<tr><td>
					<table width="100%" border="0" cellpadding="4" cellspacing="1">
						<tr>
							<td class="style1" rowspan=4 align="center" valign="top">序號</td>
							<td class="style1">建檔日期</td>
							<td class="style1">單號</td>
							<td class="style1">違規人ID</td>
							<td class="style1">違規人</td>
							<td class="style1">出生日期</td>
							<td class="style1" colspan="5">住址</td>
						</tr>
						<tr>
							<td class="style1">違規日期</td>
							<td class="style1">違規時間</td>
							<td class="style1">填單日期</td>
							<td class="style1">應到案日</td>
							<td class="style1">法條一</td>
							<td class="style1" colspan="5">違規地點</td>
						</tr>
						<tr>
							<td class="style1" colspan="2">舉發單位</td>
							<td class="style1" colspan="2">舉發員警</td>
							<td class="style1" colspan="6">到案處所</td>
						</tr>
						<tr>
							<td class="style1">裁決案號</td>
							<td class="style1">裁決日期</td>
							<td class="style1">催繳文號</td>
							<td class="style1">強制日期</td>
							<td class="style1">繳費日期</td>
							<td class="style1" colspan="2">繳費單據</td>
							<td class="style1">繳費金額</td>
							<td class="style1">繳費方式</td>
							<td class="style1">狀態</td>
						</tr>
					</table>
				</td></tr>
				<%
				set rsfound=conn.execute(strSQLTemp)
				cnt=0
				while Not rsfound.eof
					cnt=cnt+1
					'------LEO修改---------------------------------------
					ForFeit = ""
					If Not ISNull(trim(rsfound("PAYAMOUNT"))) then 
						ForFeit = trim(rsfound("PAYAMOUNT"))
					Else
						ForFeit = trim(rsfound("Forfeit1"))
					End If
					'-----------------------------------------------------
					response.write "<tr><td>"
					response.write "<table width=""100%"" border=""0"" cellpadding=""4"" cellspacing=""1"">"
					response.write "<tr>"
					response.write "<td class=""style1"" rowspan=4 align=""center"" valign=""top"">"&cnt&"&nbsp;</td>"
					response.write "<td class=""style1"" align=""left"">"&gInitDT(trim(rsfound("RecordDate")))&"</td>"
					response.write "<td class=""style1"" align=""left"">"&trim(rsfound("BillNo"))&"</td>"
					response.write "<td class=""style1"" align=""left"">"&trim(rsfound("DriverID"))&"</td>"
					response.write "<td class=""style1"" align=""left"">"&trim(rsfound("Driver"))&"</td>"
					response.write "<td class=""style1"" align=""left"">"&trim(rsfound("DriverBirth"))&"</td>"
					response.write "<td class=""style1"" align=""left"" colspan=""5"">"&trim(rsfound("DriverAddress"))&"</td>"
					response.write "</tr><tr>"
					response.write "<td class=""style1"" align=""left"">"&trim(gInitDT(DateValue(rsfound("IllegalDate"))))&"</td>"
					response.write "<td class=""style1"" align=""left"">"&trim(TimeValue(rsfound("IllegalDate")))&"</td>"
					response.write "<td class=""style1"" align=""left"">"&trim(gInitDT(rsfound("BillFillDate")))&"</td>"
					response.write "<td class=""style1"" align=""left"">"&trim(gInitDT(rsfound("DeallineDate")))&"</td>"
					response.write "<td class=""style1"" align=""left"">"&trim(rsfound("Rule1"))&"</td>"
					response.write "<td class=""style1"" align=""left"" colspan=""5"">"&trim(rsfound("IllegalAddress"))&"</td>"
					response.write "</tr><tr>"
					response.write "<td class=""style1"" align=""left"" colspan=""2"">"&trim(rsfound("BillUnitName"))&"</td>"
					response.write "<td class=""style1"" align=""left"" colspan=""2"">"&trim(rsfound("BillMem1"))&"</td>"
					response.write "<td class=""style1"" align=""left"" colspan=""6"">"&trim(rsfound("StationName"))&"</td>"
					response.write "</tr><tr>"
					response.write "<td class=""style1"" align=""left"">"&trim(rsfound("JudeNo"))&"</td>"
					response.write "<td class=""style1"" align=""left"">"&trim(gInitDT(rsfound("JudeDate")))&"</td>"
					response.write "<td class=""style1"">"&trim(rsfound("UrgeNo"))&"</td>"					
					response.write "<td class=""style1"" align=""left"">"&trim(gInitDT(rsfound("SendDate")))&"</td>"
					response.write "<td class=""style1"" align=""left"">"&trim(gInitDT(rsfound("PayDate")))&"</td>"
					response.write "<td class=""style1"" align=""left"" colspan=""2"">"&trim(rsfound("PayNo"))&"</td>"

					response.write "<td class=""style1"" align=""left"">"& ForFeit &"</td>"

					response.write "<td class=""style1"" align=""left"">"
					if trim(rsfound("PayTypeID"))="1" then
						response.write "窗口"
					elseif trim(rsfound("PayTypeID"))="2" then
						response.write "郵撥"
					end if
					response.write "&nbsp;</td>"

					response.write "<td class=""style1"" align=""left"">"
					if trim(rsfound("BILLSTATUS"))="9" then
						response.write "已繳費"
					else
						response.write "未繳費"
					end if
					response.write "&nbsp;</td>"
					response.write "</tr>"
					rsfound.MoveNext
					response.write "</table>"
					response.write "</td></tr>"
				wend
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