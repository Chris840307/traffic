<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
'On Error Resume Next
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_公示送達清冊.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950"

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

	strwhere=""
	if trim(request("Sys_SendBillSN"))<>"" then
		Sys_SendBillSN=request("Sys_SendBillSN")
		strwhere="and b.PasserSn in("&Sys_SendBillSN&")"
	else
		strwhere=request("Sys_SQL")
	end if
	orderstr=request("orderstr")

	BasSQL="select distinct a.BillNo,a.Driver,a.DriverAddress,a.illegalDate,DealLineDate,IllegalAddress,case when substr(DriverID,2,1) = 1 then '男' else '女' end as DriverSex,a.DriverBirth,c.UnitName,a.DriverID,d.illegalrule,a.Rule1,a.Forfeit1 from PassersEndArrived b,Passerbase a,UnitInfo c,Law d where b.PasserSn=a.sn and a.BillUnitID=c.UnitID and a.rule1=d.itemid and b.ArriveType=3 and d.version=2 "&strwhere

	set rs=conn.execute(BasSQL)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style1 {font-size: 12px; }
.style2 {font-size: 12px;mso-style-parent:style0;mso-number-format:"\@";}
-->
</style>
</head>
<body>
<table border=1>
 <tr>
  <td class="style1">裁決書文號</td>
  <td class="style1">受處分姓名</td>
  <td class="style1">戶籍住址</td>
  <td class="style1">違規日期</td>
  <td class="style1">違規時間</td>
  <td class="style1">原舉發單應<br>到案日期</td>
  <td class="style1">原舉通知單號</td>
  <td class="style1">違規地點</td>
  <td class="style1">性別</td>
  <td class="style1">出生年月日</td>
  <td class="style1">本分局<br>舉發所別</td>
  <td class="style1">身分證統一編號</td>
  <td class="style1">舉發違規事實</td>
  <td class="style1">違反法條</td>
  <td class="style1">裁罰金額<br>新臺幣  元</td>
  <td class="style1">公示原因</td>
 </tr>
 <%

	while Not rs.eof

		response.write "<tr>"
		response.write "<td class=""style1"">&nbsp;"&rs("BillNo")&"</td>"
		response.write "<td class=""style1"">&nbsp;"&rs("Driver")&"</td>"
		response.write "<td class=""style1"">&nbsp;"&rs("DriverAddress")&"</td>"
		response.write "<td class=""style1"">&nbsp;"&right("0"&gInitDT(rs("illegalDate")),7)&"</td>"
		response.write "<td class=""style1"">&nbsp;"&right("0"&hour(rs("illegalDate")),2)&":"&right("0"&minute(rs("illegalDate")),2)&"</td>"
		response.write "<td class=""style1"">&nbsp;"&right("0"&gInitDT(rs("DealLineDate")),7)&"</td>"
		response.write "<td class=""style1"">&nbsp;"&rs("BillNo")&"</td>"
		response.write "<td class=""style1"">&nbsp;"&rs("DriverAddress")&"</td>"
		response.write "<td class=""style1"">&nbsp;"&rs("DriverSex")&"</td>"
		If rs("DriverBirth")<>"" Then 
		response.write "<td class=""style1"">&nbsp;"&right("0"&gInitDT(rs("DriverBirth")),7)&"</td>"
		Else
		response.write "<td class=""style1"">&nbsp;</td>"
		End if
		response.write "<td class=""style1"">&nbsp;"&rs("UnitName")&"</td>"
		response.write "<td class=""style1"">&nbsp;"&rs("DriverID")&"</td>"
		response.write "<td class=""style1"">&nbsp;"&rs("illegalrule")&"</td>"
		response.write "<td class=""style1"">&nbsp;"&rs("Rule1")&"</td>"
		response.write "<td class=""style1"">&nbsp;"&rs("Forfeit1")&"</td>"
		response.write "<td class=""style1"">&nbsp;</td>"
		response.write "</tr>"
		rs.movenext
	wend
	conn.close
 %>
</table>
</body>
<script language="javascript">
window.close();
</script>
</html>