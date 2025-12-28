<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>慢車行人收據A4</title>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-family:"標楷體"; font-size: 10px; color:#ff0000;}
.style2 {font-family:"標楷體"; font-size: 10px}
.style3 {font-family:"標楷體"; font-size: 18px}
.style4 {font-family:"標楷體"; font-size: 13px}
.style5 {font-family:"標楷體"; font-size: 16px}
.style8 {font-family:"標楷體"; font-size: 36px}
.style11 {font-family:"標楷體"; font-size: 14px}
.style15 {font-family:"標楷體"; font-size: 12px}
-->
</style>
</head>

<body>
<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<%
Server.ScriptTimeout=6000

Function cmoney(istr)
	strm=split(",,拾,佰,仟,萬",",")
	strc=split("零,壹,貳,參,肆,伍,陸,柒,捌,玖",",")
	chrStr=""
	For i = len(istr) to 1 step -1

		tmpstr=mid(trim(istr),len(istr)-i+1,1)

		If i > 1 Then
			If cdbl(tmpstr) = 0 Then
				
				chkstr=mid(trim(istr),len(istr)-i+2,1)

				If cdbl(chkstr) > 0 Then

					chrStr=chrStr+strc(tmpstr)
				End if 
			else

				chrStr=chrStr+strc(tmpstr)+strm(i)
			End if 
		else

			If cdbl(tmpstr) > 0 Then chrStr=chrStr+strc(tmpstr)
		End if 
	Next

	cmoney=chrStr
End Function 

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

last_year=year(now)-1911-10

PBillSN=trim(request("PBillSN"))

strSQL="select billno,payno,payer,payamount,paydate,MIDDLEMONEY,CreditorSendNo," &_
"(select rule1 from passerBase where sn=passerpay.billsn) rule1," &_
"(select rule2 from passerBase where sn=passerpay.billsn) rule2," &_
"(select chName from MemberData where memberid=passerpay.RecordMemberID) chname," &_
"(select ReserveYear from passerBase where sn=passerpay.billsn) ReserveYear," &_
"(select Nvl(to_char(JudeDate,'YYYY'),0) JudeDate from PasserJude where BillSN=passerpay.BillSN and to_char(JudeDate,'YYYY')<>to_char(sysdate,'YYYY')) JudeDate " &_
"from passerpay where BillSN="&trim(request("PBillSN"))&" and PayTimes="&trim(request("PayTimes"))

set rspay=conn.execute(strSQL)

While Not rspay.Eof
	paydate=split(gArrDT(trim(rspay("paydate"))),"-")
%>
	<div id="L78" class="pageprint" style="position:relative;">
		<div id="Layer000" style="position:absolute; left:0px; top:0px; z-index:1"><%
				Response.Write "<img src=""./JudeJpg/PasserPay_detail02.jpg"" width=""760"">"	
			%>
		</div>

		<div id="Layer007" class="style5" style="position:absolute; left:500px; top:85px; z-index:10"><b>彰府警交字第<%=rspay("payno")%>號</b></div>

		<div id="Layer008" class="style5" style="position:absolute; left:500px; top:445px; z-index:10"><b>彰府警交字第<%=rspay("payno")%>號</b></div>

		<div id="Layer009" class="style5" style="position:absolute; left:500px; top:800px; z-index:10"><b>彰府警交字第<%=rspay("payno")%>號</b></div>



		<div id="Layer010" class="style5" style="position:absolute; left:135px; top:85px; z-index:10"><b><%=paydate(0)&"年"&paydate(1)&"月"&paydate(2)&"日"%></b></div>

		<div id="Layer010" class="style5" style="position:absolute; left:135px; top:445px; z-index:11"><b><%=paydate(0)&"年"&paydate(1)&"月"&paydate(2)&"日"%></b></div>

		<div id="Layer010" class="style5" style="position:absolute; left:135px; top:800px; z-index:12"><b><%=paydate(0)&"年"&paydate(1)&"月"&paydate(2)&"日"%></b></div>



		<div id="Layer001" class="style3" style="position:absolute; left:50px; top:170px; width:100px; z-index:10"><%=rspay("payer")%></div>

		<div id="Layer002" class="style3" style="position:absolute; left:50px; top:530px; width:100px; z-index:10"><%=rspay("payer")%></div>

		<div id="Layer003" class="style3" style="position:absolute; left:50px; top:885px; width:100px; z-index:10"><%=rspay("payer")%></div>



		<div id="Layer004" class="style3" style="position:absolute; left:275px; top:170px; width:100px; z-index:10"><%=rspay("payamount")%>元整</div>

		<div id="Layer005" class="style3" style="position:absolute; left:275px; top:525px; width:100px; z-index:10"><%=rspay("payamount")%>元整</div>

		<div id="Layer022" class="style11" style="position:absolute; left:260px; top:545px; width:100px; z-index:10"><%
			If sys_City = "彰化縣" Then 
				If rspay("JudeDate") <> "0" Then

					If trim(rspay("ReserveYear")) <> "" Then
						
						Response.Write "保留至"&rspay("ReserveYear")&"年度"

					elseif (cdbl(rspay("JudeDate"))-1911) >= last_year then 

						Response.Write "保留至"&(cdbl(rspay("JudeDate"))-1911)&"年度"
					End if 
				End if 
			end If 
		%></div>

		<div id="Layer006" class="style3" style="position:absolute; left:275px; top:880px; width:100px; z-index:10"><%=rspay("payamount")%>元整</div>

		<div id="Layer023" class="style11" style="position:absolute; left:260px; top:900px; width:100px; z-index:10"><%
			If sys_City = "彰化縣" Then 
				If rspay("JudeDate") <> "0" Then
					If trim(rspay("ReserveYear")) <> "" Then
						
						Response.Write "保留至"&rspay("ReserveYear")&"年度"

					elseif (cdbl(rspay("JudeDate"))-1911) >= last_year then 

						Response.Write "保留至"&(cdbl(rspay("JudeDate"))-1911)&"年度"
					End if 
				End if 
			end If 
		%></div>


		
		<div id="Layer011" class="style4" style="position:absolute; left:380px; top:160px; width:190px; z-index:10"><%
			Response.Write "違反道路交通管理處罰條例<br>"
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			Response.Write "第"&left(rspay("rule1"),2)&"條"
			Response.Write mid(rspay("rule1"),3,1)&"項"
			Response.Write mid(rspay("rule1"),4,2)&"款"

			If not ifnull(rspay("rule2")) Then
				
				Response.Write "<br>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Response.Write "第"&left(rspay("rule2"),2)&"條"
				Response.Write mid(rspay("rule2"),3,1)&"項"
				Response.Write mid(rspay("rule2"),4,2)&"款"
			
			End if 
		%></div>

		<div id="Layer012" class="style4" style="position:absolute; left:380px; top:520px; width:190px; z-index:10"><%
			Response.Write "違反道路交通管理處罰條例<br>"
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			Response.Write "第"&left(rspay("rule1"),2)&"條"
			Response.Write mid(rspay("rule1"),3,1)&"項"
			Response.Write mid(rspay("rule1"),4,2)&"款"

			If not ifnull(rspay("rule2")) Then
				
				Response.Write "<br>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Response.Write "第"&left(rspay("rule2"),2)&"條"
				Response.Write mid(rspay("rule2"),3,1)&"項"
				Response.Write mid(rspay("rule2"),4,2)&"款"
			
			End if 
		%></div>

		<div id="Layer013" class="style4" style="position:absolute; left:380px; top:870px; width:190px; z-index:10"><%
			Response.Write "違反道路交通管理處罰條例<br>"
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			Response.Write "第"&left(rspay("rule1"),2)&"條"
			Response.Write mid(rspay("rule1"),3,1)&"項"
			Response.Write mid(rspay("rule1"),4,2)&"款"

			If not ifnull(rspay("rule2")) Then
				
				Response.Write "<br>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Response.Write "第"&left(rspay("rule2"),2)&"條"
				Response.Write mid(rspay("rule2"),3,1)&"項"
				Response.Write mid(rspay("rule2"),4,2)&"款"
			
			End if 
		%></div>

		<div id="Layer014" class="style5" style="position:absolute; left:570px; top:185px; z-index:10"><b>字第<%=rspay("billno")%>號</b></div>

		<div id="Layer015" class="style5" style="position:absolute; left:570px; top:540px; z-index:10"><b>字第<%=rspay("billno")%>號</b></div>

		<div id="Layer016" class="style15" style="position:absolute; left:550px; top:555px; z-index:10"><b><%
			If trim(rspay("CreditorSendNo")) <>"" Then
				Response.Write "移送字第"&rspay("CreditorSendNo")&"號"
			End if 
		%></b></div>

		<div id="Layer017" class="style5" style="position:absolute; left:570px; top:895px; z-index:10"><b>字第<%=rspay("billno")%>號</b></div>

		<div id="Layer018" class="style15" style="position:absolute; left:550px; top:910px; z-index:10"><b><%
			If trim(rspay("CreditorSendNo")) <>"" Then
				Response.Write "移送字第"&rspay("CreditorSendNo")&"號"
			End if 
		%></b></div>



		<div id="Layer019" class="style3" style="position:absolute; left:190px; top:225px; z-index:10"><b><%=cmoney(rspay("payamount"))%>元整</b></div>

		<div id="Layer020" class="style3" style="position:absolute; left:190px; top:585px; z-index:10"><b><%
			Response.Write cmoney(rspay("payamount"))&"元整"
			
			If rspay("MIDDLEMONEY") > 0 Then
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Response.Write "執行必要費用新臺幣"&rspay("MIDDLEMONEY")&"元整"
			End if 
		%></b></div>

		<div id="Layer021" class="style3" style="position:absolute; left:190px; top:940px; z-index:10"><b><%
			Response.Write cmoney(rspay("payamount"))&"元整"
			
			If rspay("MIDDLEMONEY") > 0 Then
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Response.Write "執行必要費用新臺幣"&rspay("MIDDLEMONEY")&"元整"
			End if 
		%></b></div>

	</div>
<%
	rspay.MoveNext
Wend
rspay.close
set rspay=nothing
%>

</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	//window.print();
	printWindow(true,5.08,7.08,5.08,5.08);
</script>