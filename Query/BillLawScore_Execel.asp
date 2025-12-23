<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單管理</title>
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<!--#include file="sqlDCIExchangeData.asp"-->
<!-- #include file="../Common/Banner.asp"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname="績效獎勵金試算表_"&year(now)&fMnoth&fDay&".xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 

'權限
'AuthorityCheck(234)
RecordDate=split(gInitDT(date),"-")
'組成查詢SQL字串

strwhere=""
if request("IllegalDate")<>"" and request("IllegalDate1")<>""then
	ArgueDate1=gOutDT(request("IllegalDate"))&" 0:0:0"
	ArgueDate2=gOutDT(request("IllegalDate1"))&" 23:59:59"
	strwhere=" and a.IllegalDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
end if
if request("Sys_BillUnitID")<>"" then
	if strwhere<>"" then
		strwhere=strwhere&" and a.BillUnitID ='"&request("Sys_BillUnitID")&"'"
	else
		strwhere=" and a.BillUnitID='"&request("Sys_BillUnitID")&"'"
	end if
end if
if request("Sys_BillMem")<>"" then
	if strwhere<>"" then
		strwhere=strwhere&" and (a.BillMemID1='"&request("Sys_BillMem")&"' or a.BillMemID2='"&request("Sys_BillMem")&"' or a.BillMemID3='"&request("Sys_BillMem")&"')"
	else
		strwhere=" and (a.BillMemID1='"&request("Sys_BillMem")&"' or a.BillMemID2='"&request("Sys_BillMem")&"' or a.BillMemID3='"&request("Sys_BillMem")&"')"
	end if
end if
if trim(request("RecordStateID"))<>"" then
	if strwhere<>"" then
		strwhere=strwhere&" and a.RecordStateID="&request("RecordStateID")
	else
		strwhere=" and a.RecordStateID="&request("RecordStateID")
	end if
end if 

strSQL="select a.SN,a.IllegalDate,a.CarSimpleID,a.BillMem1,a.BillMem2,a.BillMem3,b.LoginID as BillMemID1,c.LoginID as BillMemID2,d.LoginID as BillMemID3,b.CreditID as CreditID1,c.CreditID as CreditID2,d.CreditID as CreditID3,b.Chname as Chname1,c.Chname as Chname2,d.Chname as Chname3,a.BillTypeID,a.BillNo,a.CarNo,a.Driver,a.DriverID,a.IllegalAddress,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.ForFeit1,a.ForFeit2,a.ForFeit3,a.ForFeit4,a.BillUnitID,a.BillStatus,a.RecordStateID,a.RecordDate,a.RecordMemberID,a.BillBaseTypeID,e.UnitName from BillBaseView a,MemberData b,MemberData c,MemberData d,UnitInfo e where a.BillMemID1=b.MemberID(+) and a.BillMemID2=c.MemberID(+) and a.BillMemID3=d.MemberID(+) and a.BillUnitID=e.UnitID(+)"&strwhere&" order by a.BillMemID1"

set rsfound=conn.execute(strSQL)
%>
<html>

</head>
<body>
<table width="100%" border="1">
	<tr>
		<td align="center"><strong>績效獎勵金試算表</strong></td>
	</tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="1">
				<tr align="center">
					<td>單位</td>
					<td>員警臂章號碼</td>
					<td>姓名</td>
					<td>身分證</td>
					<td>法條</td>
					<td>舉發單別</td>
					<td nowrap>舉發日期</td>
					<td>舉發單號</td>
				</tr>
				<tr align="center">
				<%
					while Not rsfound.eof
						chname="":chRule="":ForFeit="":CreditID="":chnameID=""
						if rsfound("BillMemID1")<>"" then chnameID=rsfound("BillMemID1")
						if rsfound("BillMemID2")<>"" then chnameID=chnameID&","&rsfound("BillMemID2")
						if rsfound("BillMemID3")<>"" then chnameID=chnameID&","&rsfound("BillMemID3")

						if rsfound("BillMemID1")<>"" then Chname=rsfound("Chname1")
						if rsfound("BillMemID2")<>"" then Chname=Chname&","&rsfound("Chname2")
						if rsfound("BillMemID3")<>"" then Chname=Chname&","&rsfound("Chname3")

						if rsfound("CreditID1")<>"" then CreditID=rsfound("CreditID1")
						if rsfound("CreditID2")<>"" then CreditID=CreditID&","&rsfound("CreditID2")
						if rsfound("CreditID3")<>"" then CreditID=CreditID&","&rsfound("CreditID3")

						if rsfound("Rule1")<>"" then chRule=rsfound("Rule1")
						if rsfound("Rule2")<>"" then chRule=chRule&"/"&rsfound("Rule2")
						if rsfound("Rule3")<>"" then chRule=chRule&"/"&rsfound("Rule3")
						if rsfound("Rule4")<>"" then chRule=chRule&"/"&rsfound("Rule4")

						response.write "<tr>"
						response.write "<td>"&rsfound("UnitName")&"&nbsp;</td>"
						response.write "<td>"&chnameID&"&nbsp;</td>"
						response.write "<td>"&Chname&"&nbsp;</td>"
						response.write "<td>"&CreditID&"&nbsp;</td>"
						response.write "<td>"&chRule&"&nbsp;</td>"
						response.write "<td>"
						if trim(rsfound("BillBaseTypeID"))="0" then
							strBTypeVal="select Content from DCIcode where TypeID=2 and ID='"&trim(rsfound("BillTypeID"))&"'"
							set rsBTypeVal=conn.execute(strBTypeVal)
							if not rsBTypeVal.eof then response.write rsBTypeVal("Content")
							rsBTypeVal.close
							set rsBTypeVal=nothing
						else
							response.write "攔停"
						end if
						response.write "&nbsp;</td>"
						response.write "<td width='5%'>"&gInitDT(trim(rsfound("IllegalDate")))&"&nbsp;</td>"
						response.write "<td width='6%'>"&rsfound("BillNo")&"&nbsp;</td>"
						response.write "</tr>"
						rsfound.movenext
					wend
				%>
				</tr>
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