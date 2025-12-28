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
fname=year(now)&fMnoth&fDay&"_慢車行人道路障礙舉發單.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 

'檢查是否可進入本系統
	strSQLTemp="select distinct a.SN,a.IllegalDate,a.BillNo,a.Driver,a.DriverID,a.DeallIneDate,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.FORFEIT1,a.FORFEIT2,a.FORFEIT3,a.FORFEIT4,a.BillStatus,a.RecordDate,a.BillFillDate,a.Note,DeCode(a.BillStatus,9,'結案',null) BillClose,a.DoubleCheckStatus,b.JUDEDATE,c.SENDDATE,c.OPENGOVNUMBER,d.URGEDATE,Decode(d.UrgeTypeID,0,'電話',1,'信函',2,'催繳書',null) UrgeTypeName,f.PayDate,f.PayNo,g.UnitName from PasserBase a,PasserJude b,PasserSend c,PasserUrge d,PassersEndArrived e,(select distinct BillSN,PayNo,PayDate from PasserPay) f,UnitInfo g where a.RecordStateID=0 and a.SN=b.BillSN(+) and a.BillNo=b.BillNo(+) and a.SN=c.BillSN(+) and a.BillNo=c.BillNo(+) and a.SN=d.BillSN(+) and a.BillNo=d.BillNo(+) and a.SN=e.PasserSN(+) and a.SN=f.BillSN(+) and a.billUnitID=g.UnitID and a.RecorDStateID<>-1"&trim(request("SQLstr"))
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>慢車行人道路障礙舉發單查詢</title>
</head>
<body>
<table width="100%" border="0">
	<tr>
		<td height="26" align="center"><strong>舉發單紀錄列表</strong></td>
	</tr>
	<tr>
		<td>
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr>
					<th>舉發類別</th>
					<th>舉發單號</th>
					<th>違規日期</th>
					<th>應到案日期</th>
					<th>罰鍰金額</th>
					<th>舉發單入案日期</th>
					<th>舉發日期</th>
					<th>舉發單狀態</th>
					<th>備註</th>
					<th>移送監理站日期</th>
					<th>裁決日期</th>
					<th>違規人姓名</th>
					<th>違規人身份證字號</th>
					<th>結案註記</th>
					<th>舉發單填單日</th>
					<th>第一次退件日期</th>
					<th>第二次郵寄日期</th>
					<th>寄存送達日期</th>
					<th>公示送達日期</th>
					<th>裁決書寄退註記</th>
					<th>催告日期</th>
					<th>繳清結案註記</th>
					<th>催繳方式</th>
					<th>裁決書寄存送達日期</th>
					<th>裁決書公示送達日期</th>
					<th>催繳通知書寄存送達日期</th>
					<th>催繳通知書公示送達日期</th>
					<th>移送執行處日期</th>
					<th>移送字號</th>
					<th>違規法條</th>
				</tr>
				<%
				BillStatusTmp=split("建檔,車籍查詢,入案,單退,寄存,公示,刪除,收受,,結案",",")
				set rsfound=conn.execute(strSQLTemp)
				while Not rsfound.eof
					response.write "<tr>"
					response.write "<td>慢車行人</td>"
					response.write "<td>"&trim(rsfound("BillNo"))&"</td>"
					response.write "<td>"&gInitDT(trim(rsfound("IllegalDate")))
					Response.Write "　"&right("00"&hour(rsfound("IllegalDate")),2)&":"&right("00"&Minute(rsfound("IllegalDate")),2)

					Response.Write "</td>"
					response.write "<td>"&gInitDT(trim(rsfound("DeallIneDate")))&"</td>"

					FORFEIT=trim(rsfound("FORFEIT1"))
					if rsfound("FORFEIT2")<>"" then FORFEIT=FORFEIT&"/"&rsfound("FORFEIT2")
					if rsfound("FORFEIT3")<>"" then FORFEIT=FORFEIT&"/"&rsfound("FORFEIT3")
					if rsfound("FORFEIT4")<>"" then FORFEIT=FORFEIT&"/"&rsfound("FORFEIT4")

					response.write "<td>"&FORFEIT&"</td>"
					response.write "<td>"&trim(gInitDT(rsfound("RecordDate")))&"</td>"
					response.write "<td>"&trim(gInitDT(rsfound("BillFillDate")))&"</td>"
					response.write "<td>"&BillStatusTmp(trim(rsfound("BillStatus")))&"</td>"
					response.write "<td>"&trim(rsfound("Note"))&"</td>"
					response.write "<td></td>"'移送監理站日期
					response.write "<td>"&trim(gInitDT(rsfound("JUDEDATE")))&"</td>"
					response.write "<td>"&trim(rsfound("Driver"))&"</td>"
					response.write "<td>"&trim(rsfound("DriverID"))&"</td>"
					response.write "<td>"&trim(rsfound("BillClose"))&"</td>"
					response.write "<td>"&trim(gInitDT(rsfound("BillFillDate")))&"</td>"
					response.write "<td></td>"
					response.write "<td></td>"
					response.write "<td></td>"
					response.write "<td></td>"
					response.write "<td></td>"
					response.write "<td>"&trim(gInitDT(rsfound("URGEDATE")))&"</td>"
					response.write "<td></td>"
					response.write "<td>"&trim(rsfound("UrgeTypeName"))&"</td>"
					response.write "<td></td>"
					response.write "<td></td>"
					response.write "<td></td>"
					response.write "<td></td>"
					response.write "<td>"&trim(gInitDT(rsfound("SENDDATE")))&"</td>"
					response.write "<td>"&trim(rsfound("OPENGOVNUMBER"))&"</td>"

					chRule=trim(rsfound("Rule1"))
					if rsfound("Rule2")<>"" then chRule=chRule&"/"&rsfound("Rule2")
					if rsfound("Rule3")<>"" then chRule=chRule&"/"&rsfound("Rule3")
					if rsfound("Rule4")<>"" then chRule=chRule&"/"&rsfound("Rule4")
					response.write "<td>"&chRule&"</td>"
					response.write "</tr>"
					rsfound.MoveNext
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