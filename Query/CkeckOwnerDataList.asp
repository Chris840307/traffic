<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_催繳車主比對清冊.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style3 {font-family:新細明體; color=0044ff; line-height:19px; font-size: 15px}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>催繳車主比對清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<%
Server.ScriptTimeout = 18000
Response.flush
'權限
'AuthorityCheck(234)
%>
<%
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	strwhere=request("SQLstr")
	
%>
</head>
<body>
<form name=myForm method="post">
<table width="100%" border="1">
	<tr>
		<td align="center" colspan="6"><strong>催繳車主比對清冊</strong></td>
	</tr>
	<tr>
		<td width="16%">單號</td>
		<td width="16%">車號</td>
		<td width="16%">停車日期</td>
		<td width="16%">催繳單號</td>
		<td width="16%">舉發單車主</td>
		<td width="16%">催繳單車主</td>
	</tr>
<%
	strSql="select x.BillNo,x.CarNo,x.imagefilename,x.OwnerA,y.imagepathname,y.Owner " &_
	",y.IllegalDate,y.ImageFileNameB " &_
	" from (select a.BillNo,a.CarNo,a.imagefilename,c.Owner as OwnerA" &_
	" from billbase a,dcilog b,billbasedcireturn c where a.sn=b.billsn and b.CarNo=c.Carno and a.RecordStateID=0 " &_
	" and b.exchangeTypeID=c.exchangeTypeID and c.exchangeTypeID='A'" &_
	" and b.Batchnumber='"&Trim(request("Sys_BatchNumber"))&"') x,BillBase y " &_
	" where  x.CarNo=y.CarNo and  x.imagefilename like '%'||y.imagepathname||'%' and y.imagepathname is not null " &_
	" and x.OwnerA<>y.Owner"
	'response.write strwhere
	Set rs1=conn.execute(strSql)
	If Not rs1.Bof Then rs1.MoveFirst 
	While Not rs1.Eof
%>
	<tr>
		<td><%=Trim(rs1("BillNO"))&"&nbsp;"%></td>
		<td><%=Trim(rs1("CarNo"))&"&nbsp;"%></td>
		<td><%=year(rs1("IllegalDate"))-1911&"/"&month(rs1("IllegalDate"))&"/"&day(rs1("IllegalDate"))&"&nbsp;"%></td>
		<td><%=Trim(rs1("ImageFileNameB"))&"&nbsp;"%></td>
		<td><%=Trim(rs1("OwnerA"))%></td>
		<td><%=Trim(rs1("Owner"))%></td>
	</tr>
<%
		
	rs1.MoveNext
	Wend
	rs1.close
	set rs1=nothing
%>
</table>
</form>
</body>
</html>
<script language="javascript">
<%conn.close%>