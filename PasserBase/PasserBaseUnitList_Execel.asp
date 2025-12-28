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
fname=year(now)&fMnoth&fDay&"_收繳費統計表.xls"
'Response.AddHeader "Content-Disposition", "filename="&fname
'response.contenttype="application/x-msexcel; charset=MS950" 

'檢查是否可進入本系統
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

	if request("PayDate1")<>"" and request("PayDate2")<>""then
		ArgueDate1=gOutDT(request("PayDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("PayDate2"))&" 23:59:59"

		paystr=" and PayDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if

	strSQLTemp="select UnitID,UnitName,Forfeit1,Count(*) as cnt,sum(nvl(Payamount,0)) Payamount from (" &_
			"		select BillUnitID UnitID,(nvl(Forfeit1,0)+nvl(Forfeit2,0)) Forfeit1," &_
			"		(select UnitName from UnitInfo where UnitID=a.BillUnitID) UnitName," &_
			"		(select Sum(NVL(Payamount,0)) Payamount from PasserPay where billsn=a.sn"&paystr&") Payamount" &_
			"		 from PasserBase a where RecordStateID=0 and Exists(select 'Y' from "&BasSQL&" where SN=a.SN) and exists(select 'Y' from PasserPay where billsn=a.sn"&paystr&")" &_
			") tmpTable Group by UnitID,UnitName,Forfeit1 order by UnitID"

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>收繳費統計表</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table width="700" border="0">
	<tr><td>
		<table width="100%" border="0" cellpadding="4" cellspacing="1">
			<tr>
				<td align="center"><strong>收繳費統計表</strong></td>
			</tr>
			<tr>
				<td align="left">列印日期：<%=gInitDt(date)%></td>
			</tr>
			<tr>
				<td align="left">處理時間：<%=request("FromILLEGALDATE")%><%if trim(request("FromILLEGALDATE"))<>"" and trim(request("TOILLEGALDATE"))<>"" then response.write "～"%><%=request("TOILLEGALDATE")%></td>
			</tr>
			<tr>
				<td align="left">登入者：<%=Session("Ch_Name")%><hr></td>
			</tr>
		</table>
	</td></tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="1">
				<tr>
					<td>舉發單位</td>
					<td>罰鍰金額</td>
					<td>舉發件數</td>
					<td>已收金額</td>
				</tr>
				<%
				set rsfound=conn.execute(strSQLTemp)
				filecnt=0:sumFile=0:sumNT=0:tmpUnitName=""
				while Not rsfound.eof
					cntNt=0
					filecnt=filecnt+cdbl(rsfound("Forfeit1"))
					sumFile=sumFile+cdbl(rsfound("cnt"))
					if trim(rsfound("Payamount"))<>"" then
						cntNt=cdbl(rsfound("Payamount"))
						sumNT=sumNT+cdbl(rsfound("Payamount"))
					end if
					response.write "<tr>"
					if trim(tmpUnitName)<>trim(rsfound("UnitName")) then
						tmpUnitName=trim(rsfound("UnitName"))
						response.write "<td align=""left"">"&trim(rsfound("UnitName"))&"</td>"
					else
						response.write "<td align=""left""></td>"
					end if
					response.write "<td align=""left"">"&trim(rsfound("Forfeit1"))&"</td>"
					response.write "<td align=""left"">"&trim(rsfound("cnt"))&"</td>"
					response.write "<td align=""left"">"&cntNt&"</td>"
					response.write "</tr>"
					rsfound.MoveNext
				wend
				response.write "<tr>"
				response.write "<td align=""right"">共計：</td>"
				response.write "<td align=""left""></td>"
				response.write "<td align=""left"">"&sumFile&"筆"&"</td>"
				response.write "<td align=""left"">"&sumNT&"元</td></tr>"
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