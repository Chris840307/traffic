<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_資料交換紀錄.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
AuthorityCheck(233)
strSQL="select distinct a.SN,a.BillNo,a.CarNo,a.Rule1,a.IllegalSpeed,a.BillMem1,c.Content,c.value,d.UnitName,e.Chname from BillBase a, BilLBaseDciReturn b, CarSpeed c,UnitInfo d,MemberData e where a.BillNo=b.BillNo and a.CarNo=b.CarNo and b.DciReturnCarType=c.ID and a.RecordMemberID=e.MemberID and a.BillUnitID=d.UnitID and a.RecordStateID=0 and b.ExChangeTypeID='W' and b.Status='Y' and a.IllegalSpeed>c.value and a.SN in(select distinct a.BillSN from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+)"&trim(request("TempSQL"))&") order by BillNo"

set rsfound=conn.execute(strSQL)
%>
<HTML>
<HEAD>
<TITLE>稽核特殊車種車速系統</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</HEAD>

<BODY>
<table width="100%" border="0">
<tr>
	<td align="center"><strong>稽核特殊車種車速列表</strong></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="4" cellspacing="1">
			<tr>
				<td>單號</td>
				<td>車號</td>
				<td>車種</td>
				<td>違規法條</td>
				<td>實際車速</td>					
				<td>設定車速</td>
				<td>舉發單位</td>
				<td>舉發員警</td>
				<td>建檔人</td>
			</tr><%
			while Not rsfound.eof
				response.write "<tr>"
				response.write "<td>"&rsfound("BillNo")&"</td>"
				response.write "<td>"&rsfound("CarNo")&"</td>"
				response.write "<td>"&rsfound("Content")&"</td>"
				response.write "<td>"&rsfound("Rule1")&"&nbsp;</td>"
				response.write "<td>"&rsfound("IllegalSpeed")&"&nbsp;</td>"
				response.write "<td>"&rsfound("Value")&"&nbsp;</td>"
				response.write "<td>"&rsfound("UnitName")&"</td>"
				response.write "<td>"&rsfound("BillMem1")&"</td>"
				response.write "<td>"&rsfound("ChName")&"</td>"
				response.write "</tr>"
				rsfound.movenext
			wend%>
		</table>
	</td>
</tr>
</table>
</BODY>
</HTML>