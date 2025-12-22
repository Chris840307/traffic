<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單庫存統計</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-size: 16px; color:#ff0000; }
.btn3{
   font-size:12px;
   font-family:新細明體;
   background-color:#EEEEEE;
   border-style:solid;
}
-->
</style>
</HEAD>
<%
	strSQL="select a.GetUnitID,(a.inNum-b.outNum-c.AccNum) totalNum from (select sum(inNum)/50 inNum,GetUnitID from (select (substr(billendnumber,3)-substr(billstartnumber,3)+1) inNum,(select unitid from unitinfo where unitname like '%交通警察%隊' and unitlevelid=1) GetUnitID from getbillbase where BillIn=1 and getbillmemberid in(select memberid from memberdata where unitid in(select unitid from unitinfo where unittypeid in(select unitid from unitinfo where unitname like '%交通警察%隊' and unitlevelid=1)))) group by GetUnitID) a,(select sum(outNum)/50 outNum,GetUnitID from (select (substr(billendnumber,3)-substr(billstartnumber,3)+1) outNum,(select unitid from unitinfo where unitname like '%交通警察%隊' and unitlevelid=1) GetUnitID from getbillbase where BillIn=2 and getbillmemberid in(select memberid from memberdata where unitid in(select unitid from unitinfo where unittypeid in(select unitid from unitinfo where unitname like '%交通警察%隊' and unitlevelid=1)))) group by GetUnitID) b,(select sum(AccNum)/50 AccNum,GetUnitID from (select (substr(billendnumber,3)-substr(billstartnumber,3)+1) AccNum,(select unitid from unitinfo where unitname like '%交通警察%隊' and unitlevelid=1) GetUnitID from getbillbase where BillIn=0 and getbillmemberid in(select memberid from memberdata where unitid in(select unitid from unitinfo where unittypeid in(select unitid from unitinfo where unitname like '%交通警察%隊' and unitlevelid=1)))) group by GetUnitID) c where a.GetUnitID=b.GetUnitID(+) and a.GetUnitID=c.GetUnitID(+)"

	strSQL=strSQL&" union all "

	strSQL=strSQL&"select a.GetUnitID,(a.inNum-b.outNum-c.AccNum) totalNum from (select sum(inNum)/50 inNum,GetUnitID from (select (substr(gtb.billendnumber,3)-substr(gtb.billstartnumber,3)+1) inNum,uit.Unittypeid GetUnitid from getbillbase gtb,memberdata mem,Unitinfo uit where gtb.BillIn=1 and gtb.getbillmemberid in(select memberid from memberdata where unitid in(select unitid from unitinfo where unittypeid in(select unitid from unitinfo where unitname like '%分局' and unitlevelid=2 and unitorder is not null))) and gtb.getbillmemberid=mem.memberid and mem.Unitid=uit.Unitid) group by GetUnitID) a,(select sum(outNum)/50 outNum,GetUnitID from (select (substr(gtb.billendnumber,3)-substr(gtb.billstartnumber,3)+1) outNum,uit.Unittypeid GetUnitid from getbillbase gtb,memberdata mem,Unitinfo uit where gtb.BillIn=2 and gtb.getbillmemberid in(select memberid from memberdata where unitid in(select unitid from unitinfo where unittypeid in(select unitid from unitinfo where unitname like '%分局' and unitlevelid=2 and unitorder is not null))) and gtb.getbillmemberid=mem.memberid and mem.Unitid=uit.Unitid) group by GetUnitID) b,(select sum(AccNum)/50 AccNum,GetUnitID from (select (substr(gtb.billendnumber,3)-substr(gtb.billstartnumber,3)+1) AccNum,uit.Unittypeid GetUnitid from getbillbase gtb,memberdata mem,Unitinfo uit where gtb.BillIn=0 and gtb.getbillmemberid in(select memberid from memberdata where unitid in(select unitid from unitinfo where unittypeid in(select unitid from unitinfo where unitname like '%分局' and unitlevelid=2 and unitorder is not null))) and gtb.getbillmemberid=mem.memberid and mem.Unitid=uit.Unitid) group by GetUnitID) c where a.GetUnitID=b.GetUnitID(+) and a.GetUnitID=c.GetUnitID(+)"

	sqlt="select uita.UnitName,geta.totalNum,uita.unitorder from ("&strSQL&") geta,UnitInfo uita where geta.getUnitID=uita.UnitID order by uita.unitorder"

	set rscnt=conn.execute(sqlt)
%>
<BODY>
<form name="myForm" method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><span class="style3">領單庫存量統計</span><b>(有做入庫的才會統計)</b></td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th class="font10">單位</th>
					<th class="font10">庫存</th>
				</tr><%
					While not rscnt.eof
						response.write "<tr bgcolor='#FFFFFF'>"
						response.write "<td class=""font10"">"
						Response.Write trim(rscnt("UnitName"))
						Response.Write "</td>"

						totalNum="<font color=""red"">資料不齊全,請確認</font>"

						If cdbl(rscnt("totalNum"))>0 Then
							totalNum=fix(rscnt("totalNum"))
						End if
						
						response.write "<td class=""font10"">"
						Response.Write totalNum
						Response.Write "</td>"

						Response.Write "<tr>"

						rscnt.movenext
					Wend
					rscnt.close
				%>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>
