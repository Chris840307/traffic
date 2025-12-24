<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_處理交通違規陳情、陳述統計表.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<!--#include virtual="traffic/Common/DB.ini"-->
<%
if request("SQLstr")<>"" then
	set rsfound=conn.execute(request("SQLstr"))
end If
Server.ScriptTimeout = 60800
Response.flush
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style1 {
	font-size:15pt; 
	font-weight: bold;
	line-height:20px;
	font-family: "標楷體";
}
.style2 {
	font-size:12pt; 
	font-family: "標楷體";
}
.style3 {
	font-size:14pt; 
	font-weight: bold;
	font-family: "標楷體";
}
.style6 {
	font-size: 16pt;
	font-weight: bold;
	line-height:20px;
	font-family: "標楷體";
}
-->
</style>
<title>處理交通違規陳情、陳述統計表</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table border="1">
	<tr>
		<td align="center" colspan="8" height="40" class="style2"><%
			strCounty="select * from apconfigure where id=35"
			Set rsC=conn.execute(strCounty)
			If Not rsC.eof Then
				response.write Trim(rsC("Value"))
			End If
			rsC.close
			Set rsC=Nothing 
		%> <%
		If Trim(request("sys_Unit"))<>"" Then
			strU="select * from UnitInfo where UnitID='"&Trim(request("sys_Unit"))&"'"
			Set rsU=conn.execute(strU)
			If Not rsU.eof Then
				response.write Trim(rsU("unitName"))
			End If 
			rsU.close
			Set rsU=Nothing 
		End If 
		%>  <%=trim(request("sys_Year"))%> 年 <%
	startMon=""
	endMon=""
	If Trim(request("sys_Date"))="1" Then
		startMon="1"
		endMon="6"
	Else
		startMon="7"
		endMon="12"
	End If 
		response.write startMon
		%> 至 <%
		response.write endMon
		%> 月處理交通違規陳情、陳述統計表</td>
	</tr>
	<tr>
		<td width="190" height="70" align="center" class="style2">項目</td>
<%	
	
	If Trim(request("sys_Unit"))<>"" Then
		sqlUnit=" and BillUnitID in (select unitId from UnitInfo where UnitTypeID='"&Trim(request("sys_Unit"))&"')"
	End If 

	For i=startMon To endMon
%>
		<td width="70" align="center" class="style2"><%=i&"月"%></td>
<%
	Next
%>
		<td width="70" align="center" class="style2">合計</td>
	</tr>
	<tr>
		<td height="80" class="style2">舉發總件數(A)</td>
<%	BillStr=""
	BillSum=0
	For i=startMon To endMon
%>
		<td align="center" class="style2"><%
		Date1=cint(request("sys_Year"))+1911&"/"&i&"/1"
		Date2=cint(request("sys_Year"))+1911&"/"&i&"/"&Day(DateAdd("d",-1,DateAdd("m",1,cint(request("sys_Year"))+1911&"/"&i&"/1")))
		strS1="select count(*) as cnt from BillbaseView where BillFillDate between to_date('"&Date1&" 0:0:0','YYYY/MM/DD HH24:MI:SS') and to_date('"&Date2&" 0:0:0','YYYY/MM/DD HH24:MI:SS') and RecordStateid=0 "&sqlUnit
		Set rs1=conn.execute(strS1)
		If Not rs1.eof Then
			response.write Trim(rs1("cnt"))
			BillSum=BillSum+CDbl(rs1("cnt"))
			If BillStr="" Then
				BillStr=Trim(rs1("cnt"))
			Else
				BillStr=BillStr&","&Trim(rs1("cnt"))
			End If 
		End If
		rs1.close 
		Set rs1=Nothing 

		'response.write Date1&"<br>"&Date2
		%></td>
<%
	Next
%>
		<td align="center" class="style2"><%
		response.write BillSum
		%></td>
	</tr>
	<tr>
		<td height="80" class="style2">公路監理機關退回更(補)正案件，民眾申訴或聲明異議等總件數(B)</td>
<%	ArgSum=0
	For i=startMon To endMon
%>
		<td align="center" class="style2"><%
		Date1=cint(request("sys_Year"))+1911&"/"&i&"/1"
		Date2=cint(request("sys_Year"))+1911&"/"&i&"/"&Day(DateAdd("d",-1,DateAdd("m",1,cint(request("sys_Year"))+1911&"/"&i&"/1")))

		strS2="select count(*) as cnt from ArgueBase a,(select Distinct BillNo,BillTypeid from billbaseView where BillNo in (select Billno from ArgueBase where Recordstateid=0 and ArgueDate between to_date('"&Date1&" 0:0:0','YYYY/MM/DD HH24:MI:SS') and to_date('"&Date2&" 0:0:0','YYYY/MM/DD HH24:MI:SS')) "&sqlUnit&") b where a.Billno=B.Billno and a.RecordStateID=0  and a.ArgueDate between to_date('"&Date1&" 0:0:0','YYYY/MM/DD HH24:MI:SS') and to_date('"&Date2&" 0:0:0','YYYY/MM/DD HH24:MI:SS')"
		Set rs2=conn.execute(strS2)
		If Not rs2.eof Then
			response.write Trim(rs2("cnt"))
			ArgSum=ArgSum+CDbl(rs2("cnt"))
		End If
		rs2.close 
		Set rs2=Nothing 
		%></td>
<%
	Next
%>
		<td align="center" class="style2"><%
		response.write ArgSum
		%></td>
	</tr>
	<tr>
		<td height="80" class="style2">上述案件(含自行查核)經查確有缺失總件數(C)</td>
<%
	BadStr=""
	BadSum=0
	For i=startMon To endMon
%>
		<td align="center" class="style2"><%
		Date1=cint(request("sys_Year"))+1911&"/"&i&"/1"
		Date2=cint(request("sys_Year"))+1911&"/"&i&"/"&Day(DateAdd("d",-1,DateAdd("m",1,cint(request("sys_Year"))+1911&"/"&i&"/1")))

		strS2="select count(*) as cnt from ArgueBase a,(select Distinct BillNo,BillTypeid from billbaseView where BillNo in (select Billno from ArgueBase where Recordstateid=0 and ArgueDate between to_date('"&Date1&" 0:0:0','YYYY/MM/DD HH24:MI:SS') and to_date('"&Date2&" 0:0:0','YYYY/MM/DD HH24:MI:SS')) "&sqlUnit&") b where a.Billno=B.Billno and a.RecordStateID=0 and a.ArgueDate between to_date('"&Date1&" 0:0:0','YYYY/MM/DD HH24:MI:SS') and to_date('"&Date2&" 0:0:0','YYYY/MM/DD HH24:MI:SS') and ((BadCnt is not null and BadCnt<>0) or (WarnCnt is not null and WarnCnt<>0)) "
		Set rs2=conn.execute(strS2)
		If Not rs2.eof Then
			response.write Trim(rs2("cnt"))
			BadSum=BadSum+CDbl(rs2("cnt"))
			If BadStr="" Then
				BadStr=Trim(rs2("cnt"))
			Else
				BadStr=BadStr&","&Trim(rs2("cnt"))
			End If 
		End If
		rs2.close 
		Set rs2=Nothing 
		%></td>
<%
	Next
%>
		<td align="center" class="style2"><%
		response.write BadSum
		%></td>
	</tr>
	<tr>
		<td height="80" class="style2">缺失比例(C/A)</td>
<%
	BillArr=Split(BillStr,",")
	BadArr=Split(BadStr,",")
	For j=0 To UBound(BillArr)
%>
		<td align="center" class="style2"><%
		If BillArr(j)="0" Then
			response.write "0"
		Else
			response.write Round(BadArr(j)/BillArr(j),5)
		End If 
		%></td>
<%
	Next
%>		<td align="center" class="style2"><%
		If BillSum<>0 then
			response.write Round(BadSum/BillSum,5)
		Else
			response.write "0"
		End If 
		%></td>
	</tr>
	<tr>
		<td height="120" align="center" class="style2">備考</td>
		<td colspan="7" class="style2">&nbsp;</td>
	</tr>

</table>
</body>
</html>
<%conn.close%>