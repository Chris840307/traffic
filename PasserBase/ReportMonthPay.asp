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

ArgueDate1=gOutDT(request("PayDate1"))&" 0:0:0"
ArgueDate2=gOutDT(request("PayDate2"))&" 23:59:59"

last_year=cdbl(left(request("PayDate2"),len(request("PayDate2"))-4))-10
now_year=cdbl(left(request("PayDate2"),len(request("PayDate2"))-4))

strwhere=strwhere&" where Exists(select 'Y' from PasserPay where PayDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and payno is not null and BillSN=PasserBase.SN) and RecordStateid=0"

sysUnit="(總表)"
whereUnit=""

if request("Sys_MemberStation")<>"" then

	strwhere=strwhere&" and MemberStation in('"&request("Sys_MemberStation")&"')"
	whereUnit=" and unitid in('"&request("Sys_MemberStation")&"')"

	strUit=split(","&replace(request("Sys_MemberStation"),"','",","),",")
end If 

Set arrCnt = Server.CreateObject("Scripting.Dictionary")

SqlUit = "select UnitID,UnitName from UnitInfo where UnitLevelID=2 and UnitName like '%分局'"&whereUnit
set rsuit=conn.execute(SqlUit)

While not rsuit.eof
	arrCnt.Add rsuit("UnitID") & "_A",""&rsuit("UnitName")&""
	arrCnt.Add rsuit("UnitID") & "_B",0
	arrCnt.Add rsuit("UnitID") & "_C",0
	arrCnt.Add rsuit("UnitID") & "_D"," "
	arrCnt.Add rsuit("UnitID") & "_E"," "
	arrCnt.Add rsuit("UnitID") & "_F",0

	For i = last_year to now_year

		arrCnt.Add i & "_"& rsuit("UnitID") &"_0",0
		arrCnt.Add i & "_"& rsuit("UnitID") &"_1",0

	Next	

	rsuit.movenext
Wend

rsuit.close

For i = last_year to now_year

	arrCnt.Add i & "_" &"A",0
	arrCnt.Add i & "_" &"B",0

Next


strSQL="select MemberStation," & _
		"(case when illegal_Year < "&last_year&" then "&now_year&" else illegal_Year end) illegal_Year," & _
		"sum(cnt) cnt,sum(PayaMount) PayaMount,min(paynomin) paynomin,max(paynomax) paynomax," & _
		"sum(paycnt) paycnt" & _
		" from (" & _
			"select MemberStation," & _
			"(case when ReserveYear is not null then to_Number(ReserveYear) " & _
			" when JudeDate is not null then to_Number(to_char(JudeDate,'YYYY'))-1911  " & _
			" else to_Number(to_char(PayDate,'YYYY'))-1911 end) illegal_Year," & _
			"sum(cnt) cnt,sum(PayaMount) PayaMount,min(paynomin) paynomin,max(paynomax) paynomax," & _
			"sum(paycnt+paydel) paycnt" & _
			" from (" & _
				"select MemberStation,ReserveYear,(select JudeDate from PasserJude where Billsn=PasserBase.sn) JudeDate," & _
				"(select max(PayDate) PayDate from PasserPay where Billsn=PasserBase.SN) PayDate," & _
				"nvl((select distinct 1 from PasserPay where CaseCloseDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and BillSN=PasserBase.SN and PasserBase.billstatus=9),0) cnt," & _
				"(select sum(nvl(PayAmount,0)) from PasserPay where PayDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and payno is not null and BillSN=PasserBase.SN) PayaMount," & _
				"(select min(payno) payno from PasserPay where PayDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and payno is not null and BillSN=PasserBase.SN) paynomin," & _
				"(select max(payno) payno from PasserPay where PayDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and payno is not null and BillSN=PasserBase.SN) paynomax," & _		
				"(select count(1) paycnt from PasserPay where PayDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and payno is not null and BillSN=PasserBase.SN) paycnt," & _
				"(select count(1) paycnt from PASSERPAYDEL where recorddate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and Billsn=PasserBase.SN) paydel" & _
				" from passerBase"& strwhere & _
			") tmbA " & _
			" group by MemberStation," & _
			" (case when ReserveYear is not null then to_Number(ReserveYear) " & _
			" when JudeDate is not null then to_Number(to_char(JudeDate,'YYYY'))-1911  " & _
			" else to_Number(to_char(PayDate,'YYYY'))-1911 end)" & _
		" ) tmpB" & _
		" group by MemberStation," & _
		" (case when illegal_Year < "&last_year&" then "&now_year&" else illegal_Year end)" & _
		" order by illegal_Year,MemberStation"

set rs=conn.execute(strSQL)

While not rs.eof
	
	arrCnt.Item(rs("illegal_Year")& "_" &rs("MemberStation")&"_0")=cdbl(rs("cnt"))
	
	arrCnt.Item(rs("illegal_Year")& "_" &rs("MemberStation")&"_1")=cdbl(rs("PayaMount"))

	arrCnt.Item(rs("MemberStation")&"_B")=arrCnt.Item(rs("MemberStation")&"_B")+cdbl(rs("cnt"))
	arrCnt.Item(rs("MemberStation")&"_C")=arrCnt.Item(rs("MemberStation")&"_C")+cdbl(rs("PayaMount"))	
	arrCnt.Item(rs("MemberStation")&"_F")=arrCnt.Item(rs("MemberStation")&"_F")+cdbl(rs("paycnt"))

	arrCnt.Item(rs("illegal_Year")&"_A")=arrCnt.Item(rs("illegal_Year")&"_A")+cdbl(rs("cnt"))
	arrCnt.Item(rs("illegal_Year")&"_B")=arrCnt.Item(rs("illegal_Year")&"_B")+cdbl(rs("PayaMount"))

	If trim(arrCnt.Item(rs("MemberStation")&"_D")) = "" Then arrCnt.Item(rs("MemberStation")&"_D")=rs("paynomin")

	If trim(arrCnt.Item(rs("MemberStation"))&"_E") = "" Then arrCnt.Item(rs("MemberStation")&"_E")=rs("paynomax")


	If arrCnt.Item(rs("MemberStation")&"_D") > rs("paynomin") Then

		arrCnt.Item(rs("MemberStation")&"_D")=rs("paynomin")
	End if 

	If arrCnt.Item(rs("MemberStation")&"_E") < rs("paynomax") Then

		arrCnt.Item(rs("MemberStation")&"_E")=rs("paynomax")
	End if 

	rs.movenext
Wend

rs.close


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>交通罰緩收入憑證月報表</title>
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

<table width="100%" border="0">
	<tr><td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td align="center" class="style1"><strong><%=thenPasserCity%></strong></td>
			</tr>
			<tr>
				<td align="center" class="style1"><strong>
				<%=rptHead1&now_year&"年"&(left(Right(request("PayDate2"),4),2))&"月份"%>違反道路交通管理事件自行繳納罰鍰明細表
				</strong></td>
			</tr>
		</table>
	</td></tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="3" cellspacing="0">
				<tr>
					<td class="style2" colspan=2>項目\單位</td>
					<%
						For i = 1 to Ubound(strUit)
							Response.Write "<td class=""style2"" width=""40"">"
							Response.Write arrCnt.Item(strUit(i)&"_A")
							Response.Write "</td>"
						Next
					%>
					<td class="style2">合計</td>
				</tr>
				<%
				For i = last_year to now_year
					Response.Write "<tr>"

					Response.Write "<td class=""style2"" rowspan=2 nowrap>"
					Response.Write i&"年度實收數"
					Response.Write "</td>"
					
					Response.Write "<td class=""style2"" nowrap>件數</td>"
					
					For j = 1 to Ubound(strUit)
						Response.Write "<td class=""style2"">"
						Response.Write arrCnt.Item(i& "_" &strUit(j)&"_0")
						Response.Write "</td>"
					next

					Response.Write "<td class=""style2"">"
					Response.Write arrCnt.Item(i&"_A")
					Response.Write "</td>"

					Response.Write "</tr>"

					Response.Write "<tr>"

					Response.Write "<td class=""style2"" nowrap>金額</td>"
					
					For j = 1 to Ubound(strUit)
						Response.Write "<td class=""style2"">"
						Response.Write arrCnt.Item(i& "_" &strUit(j)&"_1")
						Response.Write "</td>"
					next

					Response.Write "<td class=""style2"">"
					Response.Write arrCnt.Item(i&"_B")
					Response.Write "</td>"

					Response.Write "</tr>"

				next

				Response.Write "<tr>"

				Response.Write "<td class=""style2"" rowspan=2>"
				Response.Write "合計"
				Response.Write "</td>"

				Response.Write "<td class=""style2"">件數</td>"
				
				sumA=0:sumB=0
				For j = 1 to Ubound(strUit)

					sumA=sumA+cdbl(arrCnt.Item(strUit(j)&"_B"))

					Response.Write "<td class=""style2"">"
					Response.Write arrCnt.Item(strUit(j)&"_B")
					Response.Write "</td>"

				next

				Response.Write "<td class=""style2"">"
				Response.Write sumA
				Response.Write "</td>"

				Response.Write "</tr>"

				Response.Write "<tr>"

				Response.Write "<td class=""style2"">金額</td>"
				
				sumA=0:sumB=0
				For j = 1 to Ubound(strUit)

					sumB=sumB+cdbl(arrCnt.Item(strUit(j)&"_C"))

					Response.Write "<td class=""style2"">"
					Response.Write arrCnt.Item(strUit(j)&"_C")
					Response.Write "</td>"

				next

				Response.Write "<td class=""style2"">"
				Response.Write sumB
				Response.Write "</td>"

				Response.Write "</tr>"
				
				Response.Write "<tr>"

				Response.Write "<td class=""style2"" rowspan=4>本月填用或<br>註銷作癈</td>"
				Response.Write "<td class=""style2"">起號</td>"

				For j = 1 to Ubound(strUit)

					Response.Write "<td class=""style4"">"
					Response.Write arrCnt.Item(strUit(j)&"_D")
					Response.Write "&nbsp;</td>"

				next

				Response.Write "<td class=""style4"">&nbsp;</td>"
				Response.Write "</tr>"
				

				Response.Write "<tr>"
				Response.Write "<td class=""style2"">迄號</td>"

				For j = 1 to Ubound(strUit)

					Response.Write "<td class=""style4"">"
					Response.Write arrCnt.Item(strUit(j)&"_E")
					Response.Write "&nbsp;</td>"

				next

				Response.Write "<td class=""style4"">&nbsp;</td>"
				Response.Write "</tr>"

				Response.Write "<tr>"
				Response.Write "<td class=""style4"">使用張數</td>"
				
				sumC=0
				For j = 1 to Ubound(strUit)
					
					chkPayCnt=0:chkPayCnt_D=0:chkPayCnt_E=0

					If IsNumeric(right(arrCnt.Item(strUit(j)&"_D"),4)) and IsNumeric(right(arrCnt.Item(strUit(j)&"_E"),4)) Then

						chkPayCnt_D=cdbl(right(arrCnt.Item(strUit(j)&"_D"),4))
						chkPayCnt_E=cdbl(right(arrCnt.Item(strUit(j)&"_E"),4))
					End if 
					
					chkPayCnt=chkPayCnt_E-chkPayCnt_D+1

					Response.Write "<td class=""style2"">"

					If chkPayCnt_E > 0 Then
						sumC=sumC+cdbl(chkPayCnt)
						Response.Write chkPayCnt
					else
						sumC=sumC+cdbl(arrCnt.Item(strUit(j)&"_F"))
						Response.Write arrCnt.Item(strUit(j)&"_F")
					End if 
					
'					If cdbl(chkPayCnt) <> cdbl(arrCnt.Item(strUit(j)&"_F")) then 
'						Response.Write "(資料錯誤)"
'					end If 
					
					Response.Write "</td>"

				next

				Response.Write "<td class=""style2"">"&sumC&"</td>"
				Response.Write "</tr>"

				Response.Write "<tr>"
				Response.Write "<td class=""style2"">作廢</td>"

				For j = 1 to Ubound(strUit)
					strSQL="select PAYNO from PASSERPAYDEL where DELMEMBERID in(select MemberID from MemberData where UnitID in(select UnitID from Unitinfo where UnitTypeid='"&strUit(j)&"')) and PayDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"

					set rsu=conn.execute(strSQL)

					Response.Write "<td class=""style4"">"
					While not rsu.eof

						Response.Write rsu("PAYNO")&"<br>"

						rsu.movenext

					Wend

					rsu.close

					Response.Write "&nbsp;</td>"

				next

				Response.Write "<td class=""style4"">&nbsp;</td>"
				Response.Write "</tr>"

				%>
			</table>
		</td>
	</tr>
	
	<tr>
		<td class="style3">
			<!--備註：(<%=last_year%>年以前併入<%=now_year%>年度)送交通隊彙辦-->
		</td>
	</tr>
	<tr>
		<td class="style3">
			<br>
			製表&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;
			主計&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			業務&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			機關
			<br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			人員&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			組長&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			長官
		</td>
	</tr>
</table>

</body>
</html>
<%
conn.close
set conn=nothing
%>