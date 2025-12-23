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

				<td>單號</td>
				<td>舉發單位</td>
				<td>舉發人員代碼</td>
				<td>舉發人員姓名</td>
				<td>違規罰條1</td>
				<td>違規罰條2</td>
				<td>違規人身分證字號</td>
				<td>違規人</td>
				<td>違規地點</td>
				<td>違規日期</td>
				<td>填單日期</td>
				<td>建檔日期</td>
				<td>上傳日期</td>
				<td>應到案日期</td>
				</tr>
				<%
				strSql1="select a.DealLineDate,b.DciCaseInDate,a.RecordDate,a.BillFillDate,a.BillMem1,a.BillMemID1,a.BillNo,a.billtypeid,c.UnitTypeID,a.Illegaldate,a.illegaladdress " &_
					",a.Rule1,a.Rule2" &_
					",b.Driver,b.DriverID,b.DriverHomeZip,b.DriverHomeAddress,b.DriverBirthday " &_
					",b.Owner,b.OwnerID,b.OwnerZip,b.OwnerAddress " &_
					" from Billbase a,BillbaseDciReturn b,Unitinfo c " &_
					" where a.BillNo=b.BillNo and a.CarNo=b.CarNo and a.Recordstateid=0 and b.ExchangeTypeid='W' " &_
					" and a.BillUnitID=c.UnitId " &_
					" and illegaldate between to_date('"&gOutDT(request("date1"))&" 0:0:0"&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&gOutDT(request("date2"))&" 23:59:59"&"','YYYY/MM/DD/HH24/MI/SS')" &_
					" Order by c.UnitTypeID,a.BillNo"
					'response.write strSql1
				Set rsfound=conn.execute(strSql1)
					If Not rsfound.Bof Then rsfound.MoveFirst 
					While Not rsfound.Eof
						Response.flush
' 先依據舊的資料先調整了，花蓮103年審計室，如果有再調整的話，再從這個去改 by jafe 2014/01/09 16:03
'單號、舉發單位、舉發人員代碼、舉發人員姓名、違規罰條1和2、違規人身分證字號、違規地點、違規日期、填單日期、建檔日期、上傳日期、應到案日期
						response.write "<tr align='center'>"
					'單號
						response.write "<td>"
						response.write Trim(rsfound("BillNo"))
						response.write "</td>"
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
					'舉發人員代碼
						response.write "<td>"

						strUnit="select loginid from memberdata where memberid='"&Trim(rsfound("BillMemID1"))&"'"
						Set rsUnit=conn.execute(strUnit)
						If Not rsUnit.eof Then
							response.write Trim(rsUnit("loginid"))
						End If
						rsUnit.close
						Set rsUnit=Nothing 
						response.write "</td>"

					'舉發人員姓名
						response.write "<td>"
						response.write Trim(rsfound("BillMem1"))
						response.write "</td>"

					'違規罰條1
						response.write "<td>"
						response.write Trim(rsfound("Rule1"))
						response.write "</td>"
					'違規罰條2
						response.write "<td>"
						response.write Trim(rsfound("Rule2"))
						response.write "</td>"

'違規人身分證字號、違規地點、違規日期、填單日期、建檔日期、上傳日期、應到案日期
					'違規人
						response.write "<td>"
						If Trim(rsfound("billtypeid"))="1" Then
							response.write Trim(rsfound("DriverID"))
						Else
							response.write Trim(rsfound("OwnerID"))
						End If 
						response.write "</td>"
					'違規人
						response.write "<td>"
						If Trim(rsfound("billtypeid"))="1" Then
							response.write Trim(rsfound("Driver"))
						Else
							response.write Trim(rsfound("Owner"))
						End If 
						response.write "</td>"

					'違規地點
						response.write "<td>"
						response.write Trim(rsfound("IllegalAddress"))
						response.write "</td>"

					'違規時間
						response.write "<td>"
						response.write Year(rsfound("Illegaldate"))-1911&Right("00"&Month(rsfound("Illegaldate")),2)&Right("00"&day(rsfound("Illegaldate")),2)&" "&Right("00"&Hour(rsfound("Illegaldate")),2)&Right("00"&Minute(rsfound("Illegaldate")),2)
						response.write "</td>"
'違規人身分證字號、違規地點、違規日期、填單日期、建檔日期、上傳日期、應到案日期

					'填單日期
						response.write "<td>"
						response.write Year(rsfound("BillFillDate"))-1911&Right("00"&Month(rsfound("BillFillDate")),2)&Right("00"&day(rsfound("BillFillDate")),2)
						response.write "</td>"
					'建檔日期
						response.write "<td>"
						response.write Year(rsfound("RecordDate"))-1911&Right("00"&Month(rsfound("RecordDate")),2)&Right("00"&day(rsfound("RecordDate")),2)
						response.write "</td>"
					'上傳日期
						response.write "<td>"
						response.write rsfound("DciCaseInDate")
						response.write "</td>"

					'應到案日期
						response.write "<td>"
						response.write Year(rsfound("DealLineDate"))-1911&Right("00"&Month(rsfound("DealLineDate")),2)&Right("00"&day(rsfound("DealLineDate")),2)
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