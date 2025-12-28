<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsUnit=conn.execute(strSQL)
Sys_UnitID=trim(rsUnit("UnitID"))
Sys_UnitTypeID=trim(rsUnit("UnitTypeID"))
Sys_UnitLevelID=trim(rsUnit("UnitLevelID"))
rsUnit.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
elseif Sys_UnitLevelID=2 and sys_City<>"連江縣" then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
end if
set unit=conn.Execute(strSQL)
Page_UnitName=replace(trim(unit("UnitName")),"交通組","")
unit.close

strUInfo="select * from Apconfigure where ID=40"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then thenPasserCity=trim(rsUInfo("value"))
rsUInfo.close


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


strSQL="select SendDate,Driver,illegalYear,count(1) cnt from (select sn,Driver,DriverID,to_char(Illegaldate,'YYYY')-1911 illegalYear,(select to_char(min(SendDate),'YYYY')-1911 SendDate from PasserSendDetail where BillSN=PasserBase.SN group by BillSN) SendDate,(select to_char(min(PetitionDate),'YYYY')-1911 PetitionDate from PasserCreditor where BillSN=PasserBase.SN group by BillSN) PetitionDate from PasserBase where Exists(select 'Y' from "&BasSQL&" where SN=PasserBase.SN)) where SendDate is not null and PetitionDate is not null group by SendDate,Driver,illegalYear order by Driver,SendDate,illegalYear"

set rs=conn.execute(strSQL)
SendDate="":Driver="":illegalYear="":TotalNum=""
while Not rs.eof
	if trim(rs("SendDate"))<>"" then
		If not ifnull(SendDate) Then
			SendDate=SendDate&"||"
			Driver=Driver&"||"
			illegalYear=illegalYear&"||"
			TotalNum=TotalNum&"||"
		End if

		SendDate=SendDate&rs("SendDate")
		Driver=Driver&rs("Driver")
		illegalYear=illegalYear&rs("illegalYear")
		TotalNum=TotalNum&rs("cnt")
	end if
	rs.movenext
wend
rs.close

SendDate=split(SendDate,"||")
Driver=split(Driver,"||")
illegalYear=split(illegalYear,"||")
TotalNum=split(TotalNum,"||")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>債權憑證清冊</title>
<style type="text/css">
<!--
.style1 {font-size: 24px; font-family: "標楷體"; line-height:2;}
.style2 {font-size: 14px; font-family: "標楷體";}
.style3 {font-size: 16px; font-family: "標楷體";}
.style4 {font-size: 24px; font-family: "標楷體";}
-->
</style>
</head>
<body>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td class="style1" align="center">
				<%=thenPasserCity&replace(Page_UnitName,trim(thenPasserCity),"")%>
				<%=(Year(date)-1911)&"年"&Month(date)&"月份債權憑證明細統計表"%>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" border="1" cellpadding="0" cellspacing="0">
					<tr class="style3">
						<td rowspan="2" nowrap>編號</td>
						<td rowspan="2" nowrap>移送年度</td>
						<td rowspan="2" nowrap>義務人</td>
						<td nowrap><%=(Year(date)-1911-5)%>年度</td>
						<td nowrap><%=(Year(date)-1911-4)%>年度</td>
						<td nowrap><%=(Year(date)-1911-3)%>年度</td>
						<td nowrap><%=(Year(date)-1911-2)%>年度</td>
						<td nowrap><%=(Year(date)-1911-1)%>年度</td>
						<td nowrap><%=(Year(date)-1911)%>年度</td>
						<td rowspan="2" nowrap>債權件數</td>
					</tr>
					<tr>
						<td nowrap>件數</td>
						<td nowrap>件數</td>
						<td nowrap>件數</td>
						<td nowrap>件數</td>
						<td nowrap>件數</td>
						<td nowrap>件數</td>
					</tr>
					<%
					cmt=0:tmpDriver=""
					dim sumYer(6)
					For y = 0 to Ubound(sumYer)
						sumYer(y)=0
					Next
					
					For i = 0 to Ubound(Driver)
						If tmpDriver<>Driver(i) then
							tmpSendYear=""
							tmpDriver=Driver(i)
							For j = i to Ubound(Driver)
								If tmpDriver=Driver(j) and tmpSendYear<>SendDate(j) then
									tmpSendYear=SendDate(j)
									cmt=cmt+1
									SumCnt=0
									arr=-1
									Response.Write "<tr class=""style3"">"
									Response.Write "<td>"&cmt&"</td>"
									Response.Write "<td>"&SendDate(j)&"</td>"
									Response.Write "<td>"&Driver(j)&"</td>"
									
									For y = (year(date)-1911-5) to (year(date)-1911)
										chkYear=0
										arr=arr+1
										For h = j to Ubound(Driver)
											If tmpDriver=Driver(h) and tmpSendYear=SendDate(h) and trim(y)=trim(illegalYear(h)) Then
												chkYear=1
												sumYer(arr)=sumYer(arr)+cdbl(TotalNum(h))
												sumYer(6)=sumYer(6)+cdbl(TotalNum(h))
												SumCnt=SumCnt+cdbl(TotalNum(h))
												Response.Write "<td>"&TotalNum(h)&"</td>"

											elseif tmpDriver=Driver(h) and tmpSendYear<>SendDate(h) Then
												exit for

											elseif tmpDriver<>Driver(h) Then
												exit for

											End if
											
										Next
										If chkYear=0 Then Response.Write "<td>0</td>"
									Next
									Response.Write "<td>"&SumCnt&"</td></tr>"
								elseif tmpDriver<>Driver(j) then
									exit for
								end if
							Next
						end if
					next
					Response.Write "<tr>"
					Response.Write "<td colspan=""3"">合計</td>"

					For i = 0 to Ubound(sumYer)
						Response.Write "<td>"&sumYer(i)&"</td>"
					Next

					Response.Write "</tr>"
					%>
				</table>
			</td>
		</tr>
		<tr>
			<td class="style4">
				製表人：　　　　　　　　　　　組長：　　　　　　　　　　　分局長：
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
fname=year(now)&fMnoth&fDay&"_債權憑證明細統計表"
'Response.AddHeader "Content-Disposition", "filename="&fname&".xls"
'response.contenttype="application/x-msexcel; charset=MS950"
%>