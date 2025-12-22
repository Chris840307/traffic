<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
#Layer1 {
	position:absolute;
	width:209px;
	height:38px;
	z-index:2;
	top: 13px;
}
.style1 {font-size: 16px}
.style2 {font-size: 13px; }
#Layer2 {
	position:absolute;
	width:566px;
	height:38px;
	z-index:3;
	top: 15px;
}
#LayerTime {
	position:absolute;
	width:311px;
	height:34px;
	z-index:1;
}
#Layer151 {
	position:absolute;
	width:209px;
	height:38px;
	z-index:2;
}
-->
</style>
<head>
<!--#include virtual="traffic/Common/css.txt"-->
<title>領單超過30天未開單</title>
</head>
<body leftmargin="5" topmargin="5" marginwidth="0" marginheight="0">
<table width='850' border='1' align="center" >
	<tr bgcolor="#FFCC33">
		<td colspan="5"><strong>領單超過30天未開單</strong></td>
	</tr>
	<tr bgcolor="#FFFFCC">
		<td colspan="5"><strong>舉發單</strong></td>
	</tr>
	<tr>
		<td>領單時間</td>
		<td>領單單位</td>
		<td>領單人員</td>
		<td>起始單號</td>
		<td>結束單號</td>
	</tr>
<%
	strB="select getbilldate,(select (select unitname from unitinfo where unitid=memberdata.unitid) from memberdata " &_
	" where memberid=GETBILLBASE.getbillmemberid) unitname, " &_
	" (select chname from memberdata where memberid=GETBILLBASE.getbillmemberid) chname, " &_
	" billstartnumber,billendnumber " &_
	"  from GETBILLBASE where getbilldate between sysdate-180 and sysdate-30 and BillIn=0 and CounterfoiReturn=0 and not exists( " &_
	" select 'N' from billbase where billno between billstartnumber and billendnumber and recordstateid=0 " &_
	" ) and not exists( " &_
	" select 'N' from passerbase where billno between billstartnumber and billendnumber and recordstateid=0 " &_
	" )"
	Set rsB=conn.execute(strB)
	If Not rsB.Bof Then rsB.MoveFirst 
	While Not rsB.Eof
%>

	<tr>
		<td><%=rsB("getbilldate")%></td>
		<td><%=rsB("unitname")%></td>
		<td><%=rsB("chname")%></td>
		<td><%=rsB("billstartnumber")%></td>
		<td><%=rsB("billendnumber")%></td>
	</tr>
<%
		rsB.MoveNext
	Wend
	rsB.close
	set rsB=nothing
%>
<tr bgcolor="#FFFFCC">
		<td colspan="5"><strong>告示單</strong></td>
	</tr>
	<tr>
		<td>領單時間</td>
		<td>領單單位</td>
		<td>領單人員</td>
		<td>起始單號</td>
		<td>結束單號</td>
	</tr>
<%
	strB="select getbilldate, " &_
	"(select (select unitname from unitinfo where unitid=memberdata.unitid) " &_
	"from memberdata where memberid=WarningGetBillBase.getbillmemberid) unitname, " &_
	"(select chname from memberdata where memberid=WarningGetBillBase.getbillmemberid) chname, " &_
	"billstartnumber,billendnumber " &_
	" from WarningGetBillBase where getbilldate between sysdate-180 and sysdate-30 and BillIn=0 and CounterfoiReturn=0 and not exists( " &_
	"select 'N' from WarningGetBillDetail wd where exists(select 'Y' from BillReportNo where ReportNo=wd.BillNo) " &_
	"and getbillsn=WarningGetBillBase.getbillsn " &_
	")"
	Set rsB=conn.execute(strB)
	If Not rsB.Bof Then rsB.MoveFirst 
	While Not rsB.Eof
%>

	<tr>
		<td><%=rsB("getbilldate")%></td>
		<td><%=rsB("unitname")%></td>
		<td><%=rsB("chname")%></td>
		<td><%=rsB("billstartnumber")%></td>
		<td><%=rsB("billendnumber")%></td>
	</tr>
<%
		rsB.MoveNext
	Wend
	rsB.close
	set rsB=nothing
%>
<%
	conn.close
	set conn=nothing
%>	
</table>
</body>
<script type="text/javascript" src="./js/date.js"></script>
<script language="JavaScript">

</script>
</html>
