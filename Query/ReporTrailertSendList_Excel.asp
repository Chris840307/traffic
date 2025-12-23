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
<title>逕行舉發移送清冊</title>
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

tempSQL="where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and a.billno=i.billno and a.CarNo=i.CarNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and not (a.BillTypeID='2' and a.DciErrorCarData in ('1','3','9','a','j','A','H','K','T') and i.Rule4<>'2607' and a.billtypeid='2') "&Request("Sys_AllPrintSQL")&" and f.EquiPmentID=-1 and f.Note like '3%'"

strSQL="select distinct a.BillNo,f.RecordDate,f.IllegalDate,f.BillUnitID,f.DealLineDate,f.Note,f.CarNo,f.Rule1,f.Owner,f.BillFillDate from DCILog a,DCIReturnStatus d,BillBase f,(select distinct BillNo,CarNo,Rule4 from BillBaseDCIReturn where EXCHANGETYPEID='W') i "&tempSQL

strSQL=strSQL&" order by f.RecordDate"

set rsdata=conn.execute(strSQL)

strSQL="select count(1) cmt from (select distinct a.BillNo,f.IllegalDate,f.BillUnitID,f.DealLineDate,f.Note,f.CarNo,f.Rule1,f.Owner,f.BillFillDate from DCILog a,DCIReturnStatus d,BillBase f,(select distinct BillNo,CarNo,Rule4 from BillBaseDCIReturn where EXCHANGETYPEID='W') i "&tempSQL&")"

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
				<td align="center"><font size="3"><%=TitleUnitName%>&nbsp;拖吊未繳費已領單清冊</font></td>
			</tr>
			<tr>
				<td align="left">批號：<%=Request("Sys_BatchNumber")%>&nbsp; &nbsp; &nbsp; &nbsp;<%="列印日期"%>：<%=gInitDT(now)%>&nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;Page <%=fix(j/sys_cmt)+1%> of <%=fix(sys_cmt/20+0.9999999)%></td>
			</tr>
		</table>
		<table width="710" border="1" cellpadding="1" cellspacing="0">
		<tr>
		<td>
		<table width="710" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td width="5%"></td>
				<td width="10%">單號</td>
				<td width="9%">違規日期</td>
				<td width="9%"></td>
				<td width="8%"></td>
				<td width="17%"></td>
				<td width="13%">舉發單位</td>
				<td width="11%">應到案日期</td>
				<td width="11%">備註</td>
			</tr>
			<tr>
				<td>編號</td>
				<td></td>
				<td>違規時間</td>
				<td>車號</td>
				<td>法條</td>
				<td>車主</td>
				<td></td>
				<td>填單日期</td>
				<td></td>
			</tr>
		</table>
		</td>
		</tr><%
			For i = j to sys_cmt
				Response.Write "<tr><td>"
				Response.Write "<table width=""710"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
				Response.Write "<tr>"
				Response.Write "<td width=""5%"">"&i&"</td>"
				Response.Write "<td width=""10%"">"&trim(rsdata("BillNO"))&"</td>"
				Response.Write "<td width=""9%"">"&gInitDT(rsdata("IllegalDate"))&"</td>"

				Response.Write "<td width=""9%"">　</td>"
				Response.Write "<td width=""8%"">　</td>"
				Response.Write "<td width=""17%"">　</td>"

				Response.Write "<td width=""13%"">"

				strSQL="select unitname from unitinfo where unitid='"&trim(rsdata("BillUnitID"))&"'"
				set rsuit=conn.execute(strSQL)
				If not rsuit.eof Then Response.Write rsuit("unitname")
				rsuit.close

				Response.Write "</td>"

				Response.Write "<td width=""11%"">"&gInitDT(rsdata("DealLineDate"))&"</td>"
				Response.Write "<td width=""11%"">"&trim(rsdata("Note"))&"</td>"
				Response.Write "</tr>"
				Response.Write "<tr>"
				Response.Write "<td>　</td>"
				Response.Write "<td>　</td>"

				Response.Write "<td>"& Right("00"&hour(rsdata("IllegalDate")),2)&Right("00"&minute(rsdata("IllegalDate")),2)&"</td>"

				Response.Write "<td>"&trim(rsdata("CarNo"))&"</td>"
				Response.Write "<td>"&trim(rsdata("Rule1"))&"</td>"
				Response.Write "<td>"&trim(rsdata("Owner"))&"</td>"
				Response.Write "<td>　</td>"
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