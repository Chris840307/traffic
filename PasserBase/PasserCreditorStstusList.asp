<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
'if fDay<10 then	fDay="0"&fDay
'fname=year(now)&fMnoth&fDay&"_行政罰緩執行（債權）憑證處理情形統計表"
'Response.AddHeader "Content-Disposition", "filename="&fname&".xls"
'response.contenttype="application/x-msexcel; charset=MS950"

Server.ScriptTimeout = 68000
Response.flush
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

sdate="select max(illegaldate) illegaldate1,min(illegaldate) illegaldate2 from PasserBase a where a.RecordStateID=0 and Exists(select 'Y' from "&BasSQL&" where SN=a.SN)"

    set rsdate=conn.execute(sdate)

    if not rsdate.eof then
        SendDate1=split(gArrDT(rsdate("illegaldate2")),"-")
        SendDate2=split(gArrDT(rsdate("illegaldate1")),"-")
    end if

    rsdate.close


'strCnt="select count(1) cnt from (select sn,Driver,DriverID,to_char(Illegaldate,'YYYY')-1911 IllegalYear,BillNo,(select SendNumber from PasserSendDetail where SN=(select min(sn) from PasserSendDetail where BillSN=PasserBase.sn)) SendNumber,(select sum(PayAmount) from PasserPay where Billsn=PasserBase.sn) PayAmount,ForFeit1"&_
'" from PasserBase where sn in("&BaseSQL&"))a,(select BillSN,min(PetitionDate) PetitionDate from PasserCreditor group by BillSN) b where a.sn=b.billsn order by IllegalYear,DriverID,BillNo"
'
'set rscnt=conn.execute(strCnt)
'
'pagecnt=fix(cdbl(rscnt("cnt"))/30+0.9999999)
'
'rscnt.close

strSQL="select a.*,b.PetitionDate from (select sn,Driver,DriverID,to_char(Illegaldate,'YYYY')-1911 IllegalYear,BillNo,(select SendNumber from PasserSendDetail where SN=(select min(sn) from PasserSendDetail where BillSN=PasserBase.sn)) SendNumber,(select OpenGovNumber from PasserSendDetail where SN=(select min(sn) from PasserSendDetail where BillSN=PasserBase.sn)) SendOpenGovNumber,(select JudeDate from PasserJude where BillSN=PasserBase.sn) JudeDate,(select SendDate from PasserSendDetail where SN=(select min(sn) from PasserSendDetail where BillSN=PasserBase.sn)) SendDate,(select UrgeDate from PasserUrge where billsn=PasserBase.sn) UrgeDate,(select nvl(sum(PayAmount),0) from PasserPay where Billsn=PasserBase.sn) PayAmount,ForFeit1"&_
" from PasserBase where Exists(select 'Y' from "&BasSQL&" where SN=PasserBase.SN))a,(select BillSN,min(PetitionDate) PetitionDate from PasserCreditor group by BillSN) b where a.sn=b.billsn order by PetitionDate,DriverID,BillNo"

set rs=conn.execute(strSQL)

chkMonth=""
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
<%For i = 1 to 1000
	If rs.eof Then exit For 
	
	If ifnull(chkMonth) Then chkMonth=left(gInitDT(rs("PetitionDate")),5)

	If chkMonth<>left(gInitDT(rs("PetitionDate")),5)Then
		chkMonth=left(gInitDT(rs("PetitionDate")),5)
	End if 

	If i > 1 Then response.write "<div class=""PageNext"">&nbsp;</div>"
%>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td class="style1" align="center">
				<%=thenPasserCity&replace(Page_UnitName,trim(thenPasserCity),"")%>
				行政罰緩執行（債權）憑證處理情形統計表
			</td>
		</tr>
		<tr>
			<td class="style2">
				<%="　　　　　　　　　　　　　　　　　　　　　　　　　　　　　"%>
				<%=(SendDate1(0))&"年"&SendDate1(1)&"月"&SendDate1(2)&"日至"%>
				<%=(SendDate2(0))&"年"&SendDate2(1)&"月"&SendDate2(2)&"日"%>
				<%="　　　　　　　　　　　　　　　　　　　　"%>
				<%="單位：件，元"%>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" border="1" cellpadding="0" cellspacing="0">
					<tr class="style3">
						<td nowrap>取得日期</td>
						<td nowrap>編號</td>
						<td nowrap>姓名</td>
						<td nowrap>身份證字號</td>
						<td nowrap>違規年度</td>
						<td nowrap>舉發單號</td>
						<td nowrap>移送案號</td>
						<td nowrap>發文文號</td>
						<td nowrap>裁決日</td>
						<td nowrap>移送日</td>
						<td nowrap>催繳日</td>
						<td nowrap>移送金額</td>
						<td nowrap>已收繳金額</td>
						<td nowrap>再移送執行中</td>
						<td nowrap>再移送撤回核發憑證</td>
						<td nowrap>備註</td>
					</tr>
					<%
					cmt=0
					For j = 1 to 25
						If rs.eof Then exit for						
						If chkMonth<>left(gInitDT(rs("PetitionDate")),5)Then
							chkMonth=left(gInitDT(rs("PetitionDate")),5)
							exit for
						End if 
						cmt=cmt+1
						Response.Write "<tr class=""style3"">"
						Response.Write "<td>"&gInitDT(rs("PetitionDate"))&"</td>"
						Response.Write "<td>"&cmt&"</td>"
						Response.Write "<td nowrap>"&rs("Driver")&"</td>"
						Response.Write "<td>"&rs("DriverID")&"</td>"
						Response.Write "<td>"&rs("IllegalYear")&"</td>"
						Response.Write "<td nowrap>"&rs("BillNo")&"</td>"
						Response.Write "<td>"&rs("SendNumber")&"</td>"
						Response.Write "<td>"&rs("SendOpenGovNumber")&"</td>"
						Response.Write "<td>"&gInitDT(rs("JudeDate"))&"</td>"
						Response.Write "<td>"&gInitDT(rs("SendDate"))&"</td>"
						Response.Write "<td>"&gInitDT(rs("UrgeDate"))&"</td>"
						Response.Write "<td>"&rs("ForFeit1")&"</td>"
						Response.Write "<td>"&rs("PayAmount")&"</td>"
						Response.Write "<td>　</td>"
						Response.Write "<td>　</td>"
						Response.Write "<td>　</td>"
						Response.Write "</tr>"
						rs.movenext
					next

					For y=j to 25
						Response.Write "<tr class=""style3"">"
						Response.Write "<td>　</td>"
						Response.Write "<td>　</td>"
						Response.Write "<td>　</td>"
						Response.Write "<td>　</td>"
						Response.Write "<td>　</td>"
						Response.Write "<td>　</td>"
						Response.Write "<td>　</td>"
						Response.Write "<td>　</td>"
						Response.Write "<td>　</td>"
						Response.Write "<td>　</td>"
						Response.Write "<td>　</td>"
						Response.Write "<td>　</td>"
						Response.Write "<td>　</td>"
						Response.Write "<td>　</td>"
						Response.Write "<td>　</td>"
						Response.Write "<td>　</td>"
						Response.Write "</tr>"
					Next
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
<%next%>
</body>
</html>
<%
rs.close
set rsfound=nothing
conn.close
set conn=nothing


%>