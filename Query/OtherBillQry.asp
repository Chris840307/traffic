<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單另案舉發</title>
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<!--#include file="sqlDCIExchangeData.asp"-->

<%
Server.ScriptTimeout = 6800
Response.flush
'權限
'AuthorityCheck(234)
RecordDate=split(gInitDT(date),"-")

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

'組成查詢SQL字串
if request("DB_Selt")="Selt" then

end if


%>
<html>
<style type="text/css">
<!--
.style5 {
	font-size: 10pt;
}
.style7 {
	font-size: 10pt;
	font-family: "標楷體";}
.style8 {
	font-size: 14pt;
	}
.style6 {
	font-size: 16pt;
	font-weight: bold;
	line-height:22px;
	font-family: "標楷體";
}
.style11 {
	font-size: 10px;
	font-family: "標楷體";
}
.style22 {font-size: 9pt; font-family: "標楷體"; }
-->
</style>
</head>
<body>
<form name=myForm method="post">

<table width="100%" border="1" cellpadding="2" >
	<tr height="30">
		<td bgcolor="#FFCC33" colspan="9">舉發單另案舉發  <font size="4" color="red">(需刪除成功後，才會顯示另案舉發案件)</font></td>
	</tr>
	<tr bgcolor="#EBFBE3" align="center">
		<td width="12%">舉發單號</td>
		<td width="10%">車號</td>
		<td width="7%">違規日</td>
		<td width="9%">舉發員警</td>
		<td width="7%">車種</td>
		<td width="25%">違規地點</td>
		<td width="10%">法條</td>
		<td width="10%">刪除原因</td>
		<td width="10%">操作</td>
	</tr>
<%
if sys_City="苗栗縣" Then
	
	If trim(Session("Credit_ID"))<>"A01" then
		SqlPlus=" and a.Recordmemberid="&trim(Session("User_ID"))
	End if
	strsql1="select * from billbase a,BillDeleteReason b where " &_
		" a.sn not in (select Oldbillsn from otherbill)" &_
		" and a.sn=b.billsn and a.BillTypeID=2" &_
		" and a.Recordstateid=-1" & SqlPlus &_
		" and a.sn in (select billsn from ProsecutionImageDetail)" &_
		" and a.ImageFileName is not null order by a.recorddate"
else
'	strsql1="select * from billbase a,BillDeleteReason b where " &_
'		" a.sn not in (select Oldbillsn from otherbill)" &_
'		" and a.sn=b.billsn and a.BillTypeID=2" &_
'		" and b.DelReason not in ('Y','AAA','AAB','1','2','Z1','Z2','Z3') " &_
'		" and a.Recordstateid=-1 and a.Recordmemberid="&trim(Session("User_ID")) &_
'		" and a.sn in (select billsn from ProsecutionImageDetail)" &_
'		" and a.ImageFileName is not null order by a.recorddate"
	strsql1="select a.*,(select DelReason from BillDeleteReason where BillSn=a.sn) as DelReason from billbase a where " &_
		" exists (select billsn from BillDeleteReason where billsn=a.sn " &_
		" and DelReason not in ('Y','AAA','AAB','1','2','Z1','Z2','Z3') " &_
		" and DelDate>to_date('"&DATEADD("d",-15,Date())&"','yyyy/mm/dd') " &_
		" ) " &_
		" and not exists (select Oldbillsn from otherbill where Oldbillsn =a.sn ) " &_
		" and exists(select billsn from ProsecutionImageDetail where BillSn=a.sn) " &_
		" and a.BillTypeID=2 " &_
		" and a.Recorddate>to_date('"&DATEADD("d",-30,Date())&"','yyyy/mm/dd') " &_
		" and a.Recordmemberid="&trim(Session("User_ID"))&" and a.Recordstateid=-1 " &_
		" order by a.recorddate"
end if
	set rs1=conn.execute(strsql1)
	while Not rs1.eof
		isShow=0
		strE="select * from Dcilog where Billsn="&trim(rs1("Sn"))&" and ExchangeTypeID='E'"
		set rsE=conn.execute(strE)
		if rsE.eof then isShow=1
		while Not rsE.eof
			if trim(rsE("DciReturnStatusID"))="S" then isShow=1
			rsE.movenext
		wend
		rsE.close
		set rsE=nothing
	if isShow=1 then
%>
	<tr>
		<td align="center"><%=trim(rs1("BillNo"))%></td>
		<td align="center"><%=trim(rs1("CarNo"))%></td>
		<td align="center" class="style5"><%
		if not isnull(rs1("IllegalDate")) then
		response.write gInitDT(trim(rs1("IllegalDate")))&"<br>"&right("00"&hour(rs1("IllegalDate")),2)&":"&right("00"&minute(rs1("IllegalDate")),2)
		end if
		%></td>
		<td align="center"><%
		chname=""
		if rs1("BillMem1")<>"" then	chname=rs1("BillMem1")
		if rs1("BillMem2")<>"" then	chname=chname&"/"&rs1("BillMem2")
		if rs1("BillMem3")<>"" then	chname=chname&"/"&rs1("BillMem3")
		if rs1("BillMem4")<>"" then	chname=chname&"/"&rs1("BillMem4")
		response.write chname
		%></td>
		<td align="center"><%
			if trim(rs1("CarSimpleID"))="1" then
				response.write "<span class=""style5"">汽車</span>"
			elseif trim(rs1("CarSimpleID"))="2" then
				response.write "<span class=""style5"">拖車</span>"
			elseif trim(rs1("CarSimpleID"))="3" then
				response.write "<span class=""style5"">重機</span>"
			elseif trim(rs1("CarSimpleID"))="4" then
				response.write "<span class=""style5"">輕機</span>"
			elseif trim(rs1("CarSimpleID"))="6" then
				response.write "<span class=""style5"">臨時車牌</span>"
			end if
		%></td>
		<td class="style5"><%=rs1("IllegalAddress")%></td>
		<td align="center"><%
		chRule=""
		if rs1("Rule1")<>"" then chRule=rs1("Rule1")
		if rs1("Rule2")<>"" then chRule=chRule&"/"&rs1("Rule2")
		if rs1("Rule3")<>"" then chRule=chRule&"/"&rs1("Rule3")
		if rs1("Rule4")<>"" then chRule=chRule&"/"&rs1("Rule4")
		response.write chRule
		%></td>
		<td align="center"><%
		strsql2="select * from DciCode where typeid=3 and id='"&trim(rs1("DelReason"))&"'"
		set rs2=conn.execute(strsql2)
		if not rs2.eof then
			response.write rs2("Content")
		end if
		rs2.close
		set rs2=nothing
		%></td>
		<td align="center">
		<input type="button" name="save" value="另案舉發" onclick='window.open("/traffic/BillKeyIn/BillKeyIn_Car_Report.asp?BillReCover=1&ReCoverSn=<%=trim(rs1("SN"))%>&LinkUr=S","UploadFile","left=0,top=0,location=0,width=1010,height=705,resizable=yes,status=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 90px; height:26px;">
		<input type="button" name="save" value="不須另案舉發" onclick='window.open("/traffic/Query/NotOtherBill.asp?BillReCover=1&ReCoverSn=<%=trim(rs1("SN"))%>&ReCoverBillNo=<%=trim(rs1("BillNO"))%>","NotOtherBillA","left=250,top=200,location=0,width=450,height=155,resizable=yes,status=yes,scrollbars=no,menubar=no")' style="font-size: 10pt; width: 90px; height:26px;">
		</td>
	</tr>
<%	end if
		rs1.movenext
	wend
	rs1.close
	set rs1=nothing
%> 

		

</table>
<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="kinds" value="">
<input type="Hidden" name="Del_SN" value="">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">


</script>
<%
conn.close
set conn=nothing
%>