<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
Server.ScriptTimeout=12000

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

sys_City=replace(sys_City,"台中縣","台中市")
sys_City=replace(sys_City,"台南縣","台南市")

showCreditor=false
if sys_City="台中市" or sys_City = "彰化縣" or sys_City = "台南市" or sys_City = "高雄市" or sys_City = "高雄縣" or sys_City="宜蘭縣" or sys_City="基隆市" or sys_City="澎湖縣" or sys_City="屏東縣" then
	showCreditor=true
end If 

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
	strSQLTemp="select SN,CarNo,Driver,Rule1,Rule2,illegaladdress,DriverID,Driver,DriverAddress," &_
	"(select UnitName from unitinfo where unitid=a.MemberStation) StationName" &_
	" from PasserBase a where a.RecordStateID=0 and CARSIMPLEID=8 and Exists(select 'Y' from "&BasSQL&" where sn=a.sn)"&Request("orderstr")

	'If sys_City="台南市" Then ConnExecute "慢車匯出："&strSQLTemp ,360	

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style2 {font-size: 14px;mso-style-parent:style0;mso-number-format:"\@";}
-->
</style>
<title>微電車清冊</title>
</head>
<body>
<table width="100%" border="1">
	<tr>
		<td height="26" align="center"><strong>微電車清冊</strong></td>
	</tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="1">
				<tr>
					<th>編號</th>
					<th>號牌</th>
					<th>車主/負責人</th>
					<th>駕駛姓名(中/英)</th>
					<th>違規法條</th>
					<th>違規地點</th>
					<th>代保管(車輛/號牌)</th>
					<th>失竊/刑案</th>
					<th>事故紀錄</th>
					<th>公司名稱</th>
					<th>公司電話</th>
					<th>車主/公司地址</th>
					<th>居留證號</th>
					<th>居住或通訊地址</th>
					<th>聯絡電話</th>
					<th>車種</th>
					<th>廠牌</th>
					<th>國籍</th>
					<th>管轄分局(所)</th>
					<th>備註</th>					
				</tr>
				<%
				fileCnt=0
				set rsfound=conn.execute(strSQLTemp)
				while Not rsfound.eof
					fileCnt=fileCnt+1

					response.write "<tr>"
					response.write "<td>"&fileCnt& "</td>"
					response.write "<td>"&rsfound("CarNo")& "</td>"
					response.write "<td></td>"
					response.write "<td>"&trim(rsfound("Driver"))&"</td>"
					response.write "<td>"
					Response.Write trim(rsfound("Rule1"))

					If not ifnull(rsfound("Rule2")) Then
						
						Response.Write "/"&trim(rsfound("Rule1"))
					End if 
					Response.Write "</td>"

					response.write "<td>"&trim(rsfound("illegaladdress"))&"</td>"
					response.write "<td>"
					strsql="select confiscate from PASSERCONFISCATE where billsn="&rsfound("SN")
					set rs=conn.execute(strSQL)
					strConf=""
					While not rs.eof
						If strConf <> "" Then strConf=strConf&"\"

						strConf=strConf&rs("confiscate")
						rs.movenext
					Wend
					rs.close
					Response.Write strConf
					Response.Write "</td>"
					response.write "<td></td>"
					response.write "<td></td>"
					response.write "<td></td>"
					response.write "<td></td>"
					response.write "<td></td>"
					response.write "<td>"&trim(rsfound("DriverID"))&"</td>"
					response.write "<td>"&trim(rsfound("DriverAddress"))&"</td>"
					response.write "<td></td>"
					response.write "<td>微電車</td>"
					response.write "<td></td>"
					response.write "<td></td>"
					response.write "<td>"&rsfound("StationName")&"</td>"
					response.write "<td></td>"

					response.write "</tr>"

					if (fileCnt mod 10)=0 then response.flush

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

fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_微電車清冊.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>