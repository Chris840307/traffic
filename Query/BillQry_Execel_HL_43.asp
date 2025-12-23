<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_舉發單資料.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
Server.ScriptTimeout = 65000
Response.flush
%>
<%
	'檢查是否可進入本系統
	'AuthorityCheck(234)

	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單查詢</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table width="100%" border="1">
			<table width="95%" border="1" cellpadding="4" cellspacing="1">
				<tr>
					<td>單位別</td>
					<td>單號</td>
					<td>違規人</td>
					<td>出生年月日</td>
					<td>身分證字號</td>
					<td>地址</td>
					<td>違規時間</td>
					<td>違規地點</td>
					<td>違規法條</td>
					<td>車主</td>
				</tr>
				<%
				strSql1="select a.BillNo,a.billtypeid,c.UnitTypeID,a.Illegaldate,a.illegaladdress " &_
					",a.Rule1,a.Rule2" &_
					",b.Driver,b.DriverID,b.DriverHomeZip,b.DriverHomeAddress,b.DriverBirthday " &_
					",b.Owner,b.OwnerID,b.OwnerZip,b.OwnerAddress " &_
					" from Billbase a,BillbaseDciReturn b,Unitinfo c " &_
					" where a.BillNo=b.BillNo and a.CarNo=b.CarNo and a.Recordstateid=0 and b.ExchangeTypeid='W' " &_
					" and a.BillUnitID=c.UnitId and ((a.Rule1 like '431%' or a.Rule1 like '433%') or (a.Rule2 like '431%' or a.Rule2 like '433%'))" &_

					" and illegaldate between to_date('"&gOutDT(request("date1"))&" 0:0:0"&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&gOutDT(request("date2"))&" 23:59:59"&"','YYYY/MM/DD/HH24/MI/SS')" &_
					" Order by c.UnitTypeID,a.BillNo"
					'response.write strSql1
				Set rsfound=conn.execute(strSql1)
					If Not rsfound.Bof Then rsfound.MoveFirst 
					While Not rsfound.Eof
						Response.flush


						response.write "<tr align='center'>"
					'單位別
						response.write "<td>"
						'response.write Trim(rsfound("UnitTypeID"))
						strUnit="select UnitName from UnitInfo where UnitID='"&Trim(rsfound("UnitTypeID"))&"'"
						Set rsUnit=conn.execute(strUnit)
						If Not rsUnit.eof Then
							response.write Trim(rsUnit("UnitName"))
						End If
						rsUnit.close
						Set rsUnit=Nothing 
						response.write "</td>"
					'單號
						response.write "<td>"
						response.write Trim(rsfound("BillNo"))
						response.write "</td>"
					'違規人
						response.write "<td>"
						If Trim(rsfound("billtypeid"))="1" Then
							response.write Trim(rsfound("Driver"))
						Else
							response.write Trim(rsfound("Owner"))
						End If 
						response.write "</td>"
					'出生年月日
						response.write "<td>"
						response.write Trim(rsfound("DriverBirthday"))
						response.write "</td>"
					'身分證字號
						response.write "<td>"
						If Trim(rsfound("billtypeid"))="1" Then
							response.write Trim(rsfound("DriverID"))
						Else
							response.write Trim(rsfound("OwnerID"))
						End If 
						response.write "</td>"
					'地址
						response.write "<td>"
						If Trim(rsfound("billtypeid"))="1" Then
							response.write Trim(rsfound("DriverHomeZip"))&Trim(rsfound("DriverHomeAddress"))
						Else
							response.write Trim(rsfound("OwnerZip"))&Trim(rsfound("OwnerAddress"))
						End If 
						response.write "</td>"
					'違規時間
						response.write "<td>"
						response.write Year(rsfound("Illegaldate"))-1911&Right("00"&Month(rsfound("Illegaldate")),2)&Right("00"&day(rsfound("Illegaldate")),2)&" "&Right("00"&Hour(rsfound("Illegaldate")),2)&Right("00"&Minute(rsfound("Illegaldate")),2)
						response.write "</td>"
					'違規地點
						response.write "<td>"
						response.write Trim(rsfound("IllegalAddress"))
						response.write "</td>"
					'違規法條
						response.write "<td>"
						response.write Trim(rsfound("Rule1"))
						If not isnull(rsfound("Rule2")) Then
							response.write "/"&Trim(rsfound("Rule2"))
						End If 
						response.write "</td>"
					'車主
						response.write "<td>"
						response.write Trim(rsfound("Owner"))
						response.write "</td>"
						response.write "</tr>"
					rsfound.MoveNext
					Wend
					rsfound.close
					set rsfound=nothing
				%>
				</tr>
			</table>

</body>
</html>
<%
conn.close
set conn=nothing
%>