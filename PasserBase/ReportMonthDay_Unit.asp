<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%

thenPasserCity=""
strUInfo="select * from Apconfigure where ID=31"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then 
	if trim(rsUInfo("value"))<>"" and not isnull(rsUInfo("value")) then
		thenCity=replace(trim(rsUInfo("value")),"台","臺")
	end if
end if 
rsUInfo.close
set rsUInfo=nothing

sql = "select Value from Apconfigure where ID=35"
Set RSSystem = Conn.Execute(sql)
if Not RSSystem.Eof Then
	rptHead1 = RSSystem("Value")
End If 

RSSystem.close

strUit=split(",JM00,JS00,JO00,JQ00,JN00,JP00,JR00,JT00",",")

ArgueDate1=year(now)&"/"&month(now)&"/"&day(now)&" 0:0:0"
ArgueDate2=year(now)&"/"&month(now)&"/"&day(now)&" 23:59:59"

nowday=right("00"&month(now),2)&right("00"&day(now),2)
now_year=year(now)-1911

last_year=year(now)-1911-10

If trim(request("PayDate1")) <>"" and trim(request("PayDate2"))<>"" Then
	
	ArgueDate1=gOutDT(request("PayDate1"))&" 0:0:0"
	ArgueDate2=gOutDT(request("PayDate2"))&" 23:59:59"

	now_year=cdbl(left(request("PayDate2"),len(request("PayDate2"))-4))
	nowday=right(request("PayDate2"),4)
End if 


strwhere=strwhere&" where PayDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and payno is not null"

sysUnit="(總表)"
whereUnit=""

if request("Sys_MemberStation")<>"" then

	sysUnit=""

	strwhere=strwhere&" and Exists(select 'Y' from PasserBase where MemberStation in('"&request("Sys_MemberStation")&"') and sn=PasserPay.BillSN)"

	whereUnit=" and unitid='"&request("Sys_MemberStation")&"'"

end If 

Sys_Year="":str_billSN=""
Set arrCnt = Server.CreateObject("Scripting.Dictionary")

strSQL="select UitName,billSN,PayNo,billno,Driver, (case when illegalYear < "&last_year&" then "&now_year&" else illegalYear end) illegalYear,Paytatus,PayAmount,MIDDLEMONEY " & _
	" from (" & _
		"select (" & _
		"	select (" & _
		"		select UnitName from Unitinfo where Unitid=PasserBase.MemberStation" & _
		"	) uitName from passerbase where sn=passerpay.billsn" & _
		") UitName," & _	
		"billSN,PayNo," & _
		"(select billno from PasserBase where sn=passerpay.billsn) billno," & _
		"(select Driver from PasserBase where sn=passerpay.billsn) Driver," & _
		"(case when (select nvl(ReserveYear,0) from passerbase where sn=passerpay.billsn)>0 " & _
		" then (select to_Number(ReserveYear) from passerbase where sn=passerpay.billsn)" & _
		" when (select nvl(to_Number(to_char(JudeDate,'YYYY')),0) from PasserJude where billsn=passerpay.billsn)>0 " & _
		" then (select to_Number(to_char(JudeDate,'YYYY'))-1911 from PasserJude where billsn=passerpay.billsn)" & _
		" else to_Number(to_char(PayDate,'YYYY'))-1911 end) illegalYear," & _
		"(select billstatus from passerbase where sn=passerpay.billsn) Paytatus," & _
		"nvl(PayAmount,0) PayAmount,nvl(MIDDLEMONEY,0) MIDDLEMONEY " & _
		" from passerPay"& strwhere & _
	" ) PasserTmp " & _
	" order by illegalYear,PayNo"

set rs=conn.execute(strSQL)
totalcnt=0
totalMnt=0
While not rs.eof
	If instr(","&Sys_Year&",",rs("illegalYear")) <= 0 Then

		If Sys_Year <> "" Then Sys_Year=Sys_Year & ","
		Sys_Year=Sys_Year & rs("illegalYear")

		arrCnt.Add rs("illegalYear")&"_A00", ""
		arrCnt.Add rs("illegalYear")&"_A01", ""
		arrCnt.Add rs("illegalYear")&"_A02", 0
		arrCnt.Add rs("illegalYear")&"_A03", 0
		arrCnt.Add rs("illegalYear")&"_A04", 0
		arrCnt.Add rs("illegalYear")&"_A05", 0
		arrCnt.Add rs("illegalYear")&"_A09", ""
		arrCnt.Add rs("illegalYear")&"_A10", ""
		arrCnt.Add rs("illegalYear")&"_A11", ""
		arrCnt.Add rs("illegalYear")&"_A12", ""
	
	End if 
	tmpcnt=0
	
	If rs("Paytatus") = 9 Then

		If instr(","&str_billSN&",",rs("billSN")) <= 0 Then
			tmpcnt=1
		end If 
	End if 
	
	
	totalcnt=totalcnt+cdbl(tmpcnt)
	totalMnt=totalMnt+cdbl(rs("PayAmount"))+cdbl(rs("MIDDLEMONEY"))

	if request("Sys_MemberStation")<>"" then sysUnit=rs("UitName")

	arrCnt.Item(rs("illegalYear")&"_A00")=rs("UitName")
	
	arrCnt.Item(rs("illegalYear")&"_A01")=rs("illegalYear")

	arrCnt.Item(rs("illegalYear")&"_A02")=cdbl(arrCnt.Item(rs("illegalYear")&"_A02"))+cdbl(tmpcnt)
	arrCnt.Item(rs("illegalYear")&"_A03")=cdbl(arrCnt.Item(rs("illegalYear")&"_A03"))+cdbl(rs("PayAmount"))
	arrCnt.Item(rs("illegalYear")&"_A04")=cdbl(arrCnt.Item(rs("illegalYear")&"_A04"))+cdbl(rs("MIDDLEMONEY"))

	If arrCnt.Item(rs("illegalYear")&"_A09") <>"" Then 
		arrCnt.Item(rs("illegalYear")&"_A09")=arrCnt.Item(rs("illegalYear")&"_A09")&","
		arrCnt.Item(rs("illegalYear")&"_A10")=arrCnt.Item(rs("illegalYear")&"_A10")&","
	End if 

	arrCnt.Item(rs("illegalYear")&"_A09")=arrCnt.Item(rs("illegalYear")&"_A09")&rs("PayNo")
	arrCnt.Item(rs("illegalYear")&"_A10")=arrCnt.Item(rs("illegalYear")&"_A10")&rs("Driver")&"&nbsp;"&rs("BillNo")

	If cdbl(rs("MIDDLEMONEY")) > 0 Then 

		arrCnt.Item(rs("illegalYear")&"_A05")=cdbl(arrCnt.Item(rs("illegalYear")&"_A05"))+1
		
		If arrCnt.Item(rs("illegalYear")&"_A11") <>"" Then 
			arrCnt.Item(rs("illegalYear")&"_A11")=arrCnt.Item(rs("illegalYear")&"_A11")&","
			arrCnt.Item(rs("illegalYear")&"_A12")=arrCnt.Item(rs("illegalYear")&"_A12")&","
		End if 

		arrCnt.Item(rs("illegalYear")&"_A11")=arrCnt.Item(rs("illegalYear")&"_A11")&rs("PayNo")
		arrCnt.Item(rs("illegalYear")&"_A12")=arrCnt.Item(rs("illegalYear")&"_A12")&rs("Driver")&"&nbsp;"&rs("BillNo")
	end If 

	rs.movenext
Wend

rs.close


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>交通罰緩收據明細表</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<style type="text/css">
<!--
.style1 {font-size: 20px;font-family: "標楷體";line-height:2;}
.style2 {font-size: 16px;font-family: "標楷體";}
.style3 {font-size: 18px;font-family: "標楷體";}
.style4 {font-size: 10px;font-family: "標楷體";}
-->
</style>
</head>
<body>

<table width="90%" border="0">
	<tr>
		<td>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td align="center" class="style1"><strong>
					<%
						Response.Write "彰化縣警察局"&sysUnit
						Response.Write now_year&"年"&left(nowday,2)&"月"&right(nowday,2)&"日"%>收據明細表
					</strong></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="3" cellspacing="0">
				<tr>
					<td align="center" class="style2" colspan=2>分&nbsp;&nbsp;&nbsp;&nbsp;類</td>
					<td align="center" class="style2">年度</td>
					<td align="center" class="style2">件數</td>
					<td align="center" class="style2">金額</td>
					<td align="center" class="style2">收據起訖號碼</td>
					<td align="center" class="style2">備&nbsp;&nbsp;註</td>
				</tr>
				<%
				If Sys_Year <> "" Then
					Sys_Year=split(Sys_Year,",")

					For i = 0 to Ubound(Sys_Year)

						Response.Write "<tr>"

						Response.Write "<td align=""center"" class=""style2"" colspan=2 nowrap>"
						Response.Write "違反交通管理事件罰鍰"
						Response.Write "</td>"

						Response.Write "<td align=""center"" class=""style2"">"
						Response.Write arrCnt.Item(Sys_Year(i)&"_A01")
						Response.Write "</td>"

						Response.Write "<td align=""center"" class=""style2"">"
						Response.Write arrCnt.Item(Sys_Year(i)&"_A02")
						Response.Write "</td>"
						
						Response.Write "<td align=""center"" class=""style2"">"
						Response.Write arrCnt.Item(Sys_Year(i)&"_A03")
						Response.Write "</td>"

						Response.Write "<td align=""center"" class=""style2"">"
						Response.Write replace(arrCnt.Item(Sys_Year(i)&"_A09"),",","<br>")
						Response.Write "</td>"

						Response.Write "<td align=""center"" class=""style2"">"
						Response.Write replace(arrCnt.Item(Sys_Year(i)&"_A10"),",","<br>")
						Response.Write "</td>"
						
						Response.Write "</tr>"
					
					Next

					For i = 0 to Ubound(Sys_Year)

						If arrCnt.Item(Sys_Year(i)&"_A04") > 0 Then

							Response.Write "<tr>"

							Response.Write "<td align=""center"" class=""style2"" colspan=2 nowrap>"
							Response.Write "繳回年度執行手續費"
							Response.Write "</td>"

							Response.Write "<td align=""center"" class=""style2"">"
							Response.Write arrCnt.Item(Sys_Year(i)&"_A01")
							Response.Write "</td>"

							Response.Write "<td align=""center"" class=""style2"">"
							Response.Write arrCnt.Item(Sys_Year(i)&"_A05")
							Response.Write "</td>"
							
							Response.Write "<td align=""center"" class=""style2"">"
							Response.Write arrCnt.Item(Sys_Year(i)&"_A04")
							Response.Write "</td>"

							Response.Write "<td align=""center"" class=""style2"">"
							Response.Write replace(arrCnt.Item(Sys_Year(i)&"_A11"),",","<br>")
							Response.Write "</td>"
							
							Response.Write "<td align=""center"" class=""style2"">"
							Response.Write replace(arrCnt.Item(Sys_Year(i)&"_A12"),",","<br>")
							Response.Write "</td>"
							
							Response.Write "</tr>"
						end If 					
					Next
				
				End if 

				Response.Write "<tr>"

				Response.Write "<td class=""style2"" rowspan=2>"
				Response.Write "違&nbsp;&nbsp;反&nbsp;&nbsp;行&nbsp;&nbsp;政<br>"
				Response.Write "法&nbsp;&nbsp;令&nbsp;&nbsp;罰&nbsp;&nbsp;鍰"
				Response.Write "</td>"
				Response.Write "<td class=""style2"">流動人口</td>"

				Response.Write "<td class=""style2"">&nbsp;</td>"
				Response.Write "<td class=""style2"">&nbsp;</td>"
				Response.Write "<td class=""style2"">&nbsp;</td>"
				Response.Write "<td class=""style2"">&nbsp;</td>"
				Response.Write "<td class=""style2"">&nbsp;</td>"
				Response.Write "</tr>"
				
				Response.Write "<tr>"
				Response.Write "<td class=""style2"">其&nbsp;&nbsp;&nbsp;&nbsp;他</td>"

				Response.Write "<td class=""style2"">&nbsp;</td>"
				Response.Write "<td class=""style2"">&nbsp;</td>"
				Response.Write "<td class=""style2"">&nbsp;</td>"
				Response.Write "<td class=""style2"">&nbsp;</td>"
				Response.Write "<td class=""style2"">&nbsp;</td>"
				Response.Write "</tr>"

				Response.Write "<tr>"
				Response.Write "<td align=""center"" class=""style2"" colspan=2>合&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;計</td>"
				Response.Write "<td class=""style2"">&nbsp;</td>"
				Response.Write "<td class=""style2"">"&totalcnt&"</td>"
				Response.Write "<td class=""style2"">"&totalMnt&"</td>"
				Response.Write "<td class=""style2"">&nbsp;</td>"
				Response.Write "<td class=""style2"">&nbsp;</td>"
				Response.Write "</tr>"

				Response.Write "<tr>"
				Response.Write "<td align=""center"" class=""style3"">支&nbsp;票</td>"
				Response.Write "<td class=""style2"" colspan=3>新&nbsp;臺&nbsp;幣&nbsp;&nbsp;&nbsp;&nbsp;</td>"
				Response.Write "<td class=""style3"" colspan=3>支票號碼：</td>"
				Response.Write "</tr>"

				Response.Write "<tr>"
				Response.Write "<td align=""center"" class=""style3"">現&nbsp;金</td>"
				Response.Write "<td class=""style2"" colspan=3>新&nbsp;臺&nbsp;幣&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;元</td>"
				Response.Write "<td class=""style3"" colspan=3>出納&nbsp;&nbsp;&nbsp;年&nbsp;&nbsp;&nbsp;月&nbsp;&nbsp;&nbsp;日簽收：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>"
				Response.Write "</tr>"

				%>
			</table>
		</td>
	</tr>
	<tr>
		<td class="style3">
			<br>
			製表&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;
			主管&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;
			會計&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;
			<br><br>
			總局出納&nbsp;&nbsp;&nbsp;年&nbsp;&nbsp;&nbsp;月&nbsp;&nbsp;&nbsp;日簽收 ：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</td>
	</tr>
</table>

</body>
</html>
<%
conn.close
set conn=nothing
%>