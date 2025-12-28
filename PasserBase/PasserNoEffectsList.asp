<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

Server.ScriptTimeout = 16000
Response.flush
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style3 {font-family:新細明體; color=0044ff; line-height:19px; font-size: 15px}
.style4 {font-family:新細明體; color=0044ff; line-height:12px; font-size: 10px}
.style5 {font-family:新細明體; color=0044ff; line-height:13px; font-size: 11px}
.style6 {font-family:新細明體; color=0044ff; line-height:12px; font-size: 10px}
-->
</style>
<style media="print">
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>無個人財產清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<%

strUnitName="select Value from ApConfigure where ID=40"
set rsUnitName=conn.execute(strUnitName)
if not rsUnitName.eof then
	TitleUnitName=trim(rsUnitName("value"))&" "&TitleUnitName2
end if
rsUnitName.close
set rsUnitName=nothing

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

strSQL="select a.SN,a.IllegalDate,a.BillNo,a.Driver,a.DriverID,a.DriverAddress," & _
		"a.Rule1,a.BillUnitID,a.BillFillDate,a.DealLineDate,a.Note" & _
		" from PasserBase a where a.RecordStateID=0 and Exists(select 'Y' from "&BasSQL&" where SN=a.SN) and exists(select 'Y' from PasserSendDetail where a.sn=BillSn and exists(select 'Y' from PasserCreditor where CreditorTypeID='1' and SendDetailSN=PasserSendDetail.sn))"&Request("orderstr")

set rsdata=conn.execute(strSQL)

strSQL="select count(1) cmt from PasserBase a where a.RecordStateID=0 and Exists(select 'Y' from "&BasSQL&" where SN=a.SN) and exists(select 'Y' from PasserSendDetail where a.sn=BillSn and exists(select 'Y' from PasserCreditor where CreditorTypeID='1' and SendDetailSN=PasserSendDetail.sn))"

set rscmt=conn.execute(strSQL)
sys_cmt=cdbl(rscmt("cmt"))
rscmt.close

If rsdata.eof Then Response.End
%>
</head>
<body>
<object id="factory" style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://10.104.10.246/traffic/smsx.cab#Version=6,1,432,1">
</object>
<form name=myForm method="post"><%
	For j = 1 to sys_cmt step 20
		If j>1 Then response.write "<div class=""PageNext"">　</div>"%>
		<table width="710" border="0" cellpadding="1" cellspacing="0">
			<tr>
				<td align="center"><font size="3"><%=TitleUnitName%>&nbsp;無個人財產清冊</font></td>
			</tr>
			<tr>
				<td align="right"><%="列印日期"%>：<%=gInitDT(now)%>&nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;Page <%=fix(j/20+0.9999999)%> of <%=fix(sys_cmt/20+0.9999999)%></td>
			</tr>
		</table>
		<table width="710" border="1" cellpadding="1" cellspacing="0">
		<tr>
		<td>
		<table width="710" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td width="5%" rowspan="2">編號</td>
				<td width="10%" rowspan="2">單號</td>
				<td width="9%">違規日期</td>
				<td width="9%">違規人</td>
				<td width="20%" rowspan="2">違規人地址</td>
				<td width="15%">舉發單位</td>
				<td width="11%">應到案日期</td>
				<td width="11%" rowspan="2">備註</td>
			</tr>
			<tr>
				<td>違規時間</td>
				<td>違規人證號</td>
				<td>法條</td>
				<td>填單日期</td>
			</tr>
		</table>
		</td>
		</tr><%
			For i = j to sys_cmt
				Response.Write "<tr><td>"
				Response.Write "<table width=""710"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
				Response.Write "<tr>"
				Response.Write "<td width=""5%"" rowspan=""2"">"&i&"</td>"

				Response.Write "<td width=""10%""  rowspan=""2"">"&trim(rsdata("BillNo"))&"</td>"

				Response.Write "<td width=""9%"">"&gInitDT(rsdata("IllegalDate"))&"</td>"

				Response.Write "<td width=""9%"">"&trim(rsdata("Driver"))&"</td>"

				Response.Write "<td width=""20%""  rowspan=""2"">"&trim(rsdata("DriverAddress"))&"</td>"

				Response.Write "<td width=""15%"">"

				strSQL="select unitname from unitinfo where unitid='"&trim(rsdata("BillUnitID"))&"'"
				set rsuit=conn.execute(strSQL)
				If not rsuit.eof Then Response.Write rsuit("unitname")
				rsuit.close

				Response.Write "</td>"

				Response.Write "<td width=""11%"">"&gInitDT(rsdata("DealLineDate"))&"</td>"

				Response.Write "<td width=""11%"" rowspan=""2"">"&trim(rsdata("Note"))&"</td>"

				Response.Write "</tr>"
				Response.Write "<tr>"

				Response.Write "<td>"& Right("00"&hour(rsdata("IllegalDate")),2)&Right("00"&minute(rsdata("IllegalDate")),2)&"</td>"

				Response.Write "<td>"&trim(rsdata("DriverID"))&"</td>"
				Response.Write "<td>"&trim(rsdata("Rule1"))&"</td>"
				Response.Write "<td>"&gInitDT(rsdata("BillFillDate"))&"</td>"
				Response.Write "</tr>"
				Response.Write "</table>"
				Response.Write "</td>"
				Response.Write "</tr>"
				rsdata.movenext
				If i-j = 19 then exit for					
			Next%>
		</table>
	<%next
	rsdata.close%>
</form>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	printWindow(true,7,5.08,5.08,5.08);
</script>