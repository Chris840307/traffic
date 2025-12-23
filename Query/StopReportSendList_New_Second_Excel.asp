<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>催繳資料清冊</title>
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<style type="text/css">
<!--
.style1 {
	font-size: 14pt;
	line-height: 20px;
}
.style2 {
	font-size: 11pt;
}

-->
</style>
<%
Server.ScriptTimeout = 800
Response.flush
'權限
'AuthorityCheck(234)

RecordDate=split(gInitDT(date),"-")

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

	PageCount=21
	
	Sys_SendMarkDate1=gOutDT(request("Sys_SendMarkDate1"))&" 0:0:0"
	Sys_SendMarkDate2=gOutDT(request("Sys_SendMarkDate2"))&" 23:59:59"

	strCnt="select count(*) as cnt from (select ImageFileNameB,Max(SN) as SN" &_
	" from Billbase  where RecordStateID=0 group by ImageFileNameB) k,BillBase a,stopcarsendaddress b where a.ImageFileNameB=b.BillNo and a.RecordStateID=0 and k.SN=a.SN and b.UserMarkDate between TO_DATE('"&Sys_SendMarkDate1&"','YYYY/MM/DD/HH24/MI/SS') " &_
	" and TO_DATE('"&Sys_SendMarkDate2&"','YYYY/MM/DD/HH24/MI/SS')" 
	'response.write strCnt
	'response.end
	set Dbrs=conn.execute(strCnt)
	DBsum=Dbrs("cnt")
	if trim(Dbrs("cnt"))="0" then
		pagecnt=1
	else
		pagecnt=fix(Cint(Dbrs("cnt"))/PageCount+0.9999999)
	end if
	Dbrs.close


	strSQL="select distinct a.* from (select ImageFileNameB,Max(SN) as SN" &_
	" from Billbase  where RecordStateID=0 group by ImageFileNameB) k,BillBase a,stopcarsendaddress b where a.ImageFileNameB=b.BillNo and a.RecordStateID=0 and k.SN=a.SN and b.UserMarkDate between TO_DATE('"&Sys_SendMarkDate1&"','YYYY/MM/DD/HH24/MI/SS') " &_
	" and TO_DATE('"&Sys_SendMarkDate2&"','YYYY/MM/DD/HH24/MI/SS') order by a.RecordDate"
	set rsfound=conn.execute(strSQL)
'response.write strwhere
%>

</head>
<body>
<form name=myForm method="post">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;


承辦人：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				單位主管：
<%
mailSN=0
If Not rsfound.Bof Then rsfound.MoveFirst 
While Not rsfound.Eof
if mailSN>0 then response.write "<div class=""PageNext"">&nbsp;</div>"
%>

	<table width="100%" border="0" cellspacing="0">
		<tr>
			<td height="10"></td>
		</tr>
		<tr>
		<td width="40%">
		<span class="style2">
		列印日期：<%=now%><br>
		列印單位：<%
		strPrintUnit="select UnitName from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"'"
		set rsPrintUnit=conn.execute(strPrintUnit)
		if not rsPrintUnit.eof then
			response.write rsPrintUnit("UnitName")
		end if
		rsPrintUnit.close
		set rsPrintUnit=nothing
		%><br>
		列印人員：<%
		strUserID="select Chname from MemberData where MemberID='"&trim(Session("User_ID"))&"'"
		set rsUserID=conn.execute(strUserID)
		if not rsUserID.eof then
			response.write rsUserID("ChName")
		end if
		rsUserID.close
		set rsUserID=nothing
		%>
		</span>
		</td>
		<td width="30%">
		<span class="style1">催繳資料清冊</span>		
		</td>
		<td width="30%" align="right">
		<span class="style2">
		頁次 <%=fix(mailSN/PageCount)+1%> of <%=pagecnt%> &nbsp; &nbsp; 
		</span>
		</td>
		</tr>
	</table>
	<br>
	<table width="100%" border="0" cellspacing="0">
		<tr>
		<td width="14%"><span class="style2">停車日期時間</span></td>
		<td width="15%"><span class="style2">車號</span></td>
		<td width="56%"><span class="style2">停車地點</span></td>
		<td width="15%"><span class="style2">應繳停車費</span></td>
		</tr>
		<tr>
		<td><span class="style2">車種</span></td>
		<td><span class="style2">車主姓名</span></td>
		<td><span class="style2">地址</span></td>
		<td></td>
		</tr>
	</table>
	<hr>
<%		for i=1 to PageCount
			if rsfound.eof then exit for
			mailSN=mailSN+1
%>
	<table width="100%" border="0" cellspacing="0">
	<tr>
		<td width="14%"><span class="style2"><%
		response.write year(rsfound("IllegalDate"))-1911&"/"& month(rsfound("IllegalDate"))& "/" &day(rsfound("IllegalDate"))
		response.write "&nbsp; "&right("00"&hour(rsfound("IllegalDate")),2)&":"&right("00"&minute(rsfound("IllegalDate")),2)
		%></span></td>
		<td width="15%"><span class="style2"><%
		response.write rsfound("CarNo")
		%></span></td>
		<td width="56%"><span class="style2"><%
		response.write funcCheckFont(trim(rsfound("IllegalAddress")),17,1)		
		%></span></td>
		<td width="15%"><span class="style2"><%
		response.write trim(rsfound("Forfeit1"))		
		%></span></td>
	</tr>
	<tr>
		<td><span class="style2">
		<%
		if trim(rsfound("CarSimpleID"))="1" then
			response.write "汽車"
		elseif trim(rsfound("CarSimpleID"))="2" then
			response.write "拖車"
		elseif trim(rsfound("CarSimpleID"))="3" then
			response.write "重機"
		elseif trim(rsfound("CarSimpleID"))="4" then
			response.write "輕機"
		elseif trim(rsfound("CarSimpleID"))="6" then
			response.write "臨時車牌"
		end if
		%>
		</span></td>
		<td><span class="style2">
		<%
		If Not isNull(rsfound("Owner")) then
			If InStr(trim(rsfound("Owner")),"@@")>0 Then
				response.write funcCheckFont(trim(rsfound("Owner")),17,1)
			else
				response.write funcCheckFont(left(trim(rsfound("Owner")),7),17,1)
			End If
		End if
		%>
		</span></td>
		<td><span class="style2">
		<%
		if trim(rsfound("DriverAddress"))<>"" and not isnull(rsfound("DriverAddress")) then
			response.write trim(rsfound("DriverZip"))&funcCheckFont(trim(rsfound("DriverAddress")),17,1)
		end if
		%>
		</span></td>
		
	</tr>
		<tr>
			<td height="7"></td>
		</tr>
	</table>
<%			
		rsfound.MoveNext
		next
%>

<%
Wend
rsfound.close
set rsfound=nothing
%>
</form>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
//printWindow(true,7,5.08,5.08,5.08);
</script>
<%conn.close%>