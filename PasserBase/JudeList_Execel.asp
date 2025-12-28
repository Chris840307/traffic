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
fname=year(now)&fMnoth&fDay&"_裁決清冊.xls"
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

	strSQLTemp="select SN,IllegalDate,BillNo,Driver,DriverID,Rule1,BillFillDate,DeallineDate,FORFEIT1," &_
	"(Select JudeDate from PasserJude where billsn=PasserBase.sn) JUDEDATE," &_
	"(Select OpenGovNumber from PasserJude where billsn=PasserBase.sn) OpenGovNumber," &_
	"(Select UnitName from UnitInfo where UnitID=PasserBase.BillUnitID) BillUnitName" &_
	" from PasserBase where RecordStateID=0 and Exists(select 'Y' from "&BasSQL&" where sn=PasserBase.sn) and Exists(select 'Y' from PasserJude where Billsn=PasserBase.SN) order by DriverID"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>裁決清冊</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table width="700" border="1">
	<tr><td>
		<table width="100%" border="0" cellpadding="4" cellspacing="1">
			<tr>
				<td align="center"><strong>裁決清冊</strong></td>
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

'				strSQL="select sum(Payamount) as Sys_Payamount from PasserPay where BillSN="&rsfound("SN")&" and BillNo='"&rsfound("BillNo")&"'"
'				set rspay=conn.execute(strSQL)
'				if not rspay.eof then Sys_Payamount=rspay("Sys_Payamount")
'				rspay.close
'
				if Not rsfound.eof then
					tmpDriverID=trim(rsfound("DriverID"))
					showDriverID=trim(rsfound("DriverID"))
					showDriver=trim(rsfound("Driver"))
					'if Not isnull(Sys_Payamount) then sumNT=Cint(Sys_Payamount)
				end if
				response.write "<tr><td>"
				response.write "<table width=""100%"" border=""0"" cellpadding=""4"" cellspacing=""1"">"
				while Not rsfound.eof
					sumFile=sumFile+1
'					Sys_Payamount=0
'					strSQL="select sum(Payamount) as Sys_Payamount from PasserPay where BillSN="&rsfound("SN")&" and BillNo='"&rsfound("BillNo")&"'"
'					set rspay=conn.execute(strSQL)
'					if not rspay.eof then
						tmpDriverID=trim(rsfound("DriverID"))
						Sys_Payamount=rsfound("FORFEIT1")
'						if Not isnull(rspay("Sys_Payamount")) then Sys_Payamount=rspay("Sys_Payamount")
'					end if
'					rspay.close

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
					response.write "<td align=""left"">"&trim(gInitDT(rsfound("JudeDATE")))
					if trim(rsfound("OpenGovNumber"))<>"" then response.write "/"&trim(rsfound("OpenGovNumber"))
					response.write "</td>"
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