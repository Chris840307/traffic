<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單列印-Legal Size</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-size: 9px}
.style2 {font-size: 10px}
.style3 {font-size: 14px}
.style4 {font-size: 12px}
.style7 {font-size: 13px}
.style8 {font-size: 36px}
.style11 {font-size: 14px}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
</head>

<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<%
PBillSN=split(trim(request("PBillSN")),",")
Server.ScriptTimeout=600
Sys_IMAGEFILENAME=split(",",",")
Sys_IMAGEFILENAMEB=split(",",",")
Sys_IMAGEPATHNAME=split(",",",")
Sys_CarNo=split(",",",")
BillNo=split(",",",")
for i=0 to Ubound(PBillSN) step 2
	if cint(i)<>0 then response.write "<div class=""PageNext""></div>"
	for j=0 to 1
		sumCnt=sumCnt+1
		if j>0 and i>=Ubound(PBillSN) then exit for
		strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i+j)
		set rsbil=conn.execute(strBil)
		BillNo(j)=trim(rsbil("BillNo"))
		Sys_CarNo(j)=trim(rsbil("CarNo"))
		strSql="select distinct BillFillerMemberID,DriverSex,DriverBirth,IllegalDate,DealLineDate,IMAGEFILENAME,IMAGEPATHNAME,IMAGEFILENAMEB from BillBase where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"'"
		set rssex=conn.execute(strSql)
		Sys_IMAGEFILENAME(j)="":Sys_IMAGEFILENAMEB(j)=""
		if Not rssex.eof then Sys_IMAGEFILENAME(j)=trim(rssex("IMAGEFILENAME"))
		if Not rssex.eof then Sys_IMAGEFILENAMEB(j)=trim(rssex("IMAGEFILENAMEB"))
		if Not rssex.eof then Sys_IMAGEPATHNAME(j)=trim(rssex("IMAGEPATHNAME"))
		rssex.close
		rsbil.close
		if err.Number<>0 then msgBox("資料讀取錯誤"&Cstr(Err.Number)&Err.Description)
		err.Clear
		%>
		<!--<div style="position:absolute; left:10px; top:<%=10+406*(i+j)%>px;">-->
		<table width="645" height="500" border="0">
			<tr>
				<td colspan=2><font size="3"><%=BillNo(j)+"，"+Sys_CarNo(j)%></font></td>
			</tr>
			<tr>
				<td width="10%" valign="bottom">
					<%if trim(Sys_IMAGEFILENAMEB(j))<>"" then%>
						<img src=<%=""""&Sys_IMAGEPATHNAME(j)&Sys_IMAGEFILENAMEB(j)&""""%> width="210" height="165">
					<%end if%>
				</td>
				<td width="70%" valign="bottom">
					<%if trim(Sys_IMAGEFILENAME(j))<>"" then%>
						<img src=<%=""""&Sys_IMAGEPATHNAME(j)&Sys_IMAGEFILENAME(j)&""""%> width="565" height="365">
					<%end if%>
				</td>
			</tr>
		</table>
		<!--</div>-->
	<%next
next%>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
</script>