<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_受理民眾交通違規申訴項目分析及員警錯誤舉發與撤銷案件統計表.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=Nothing

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
	font-weight: bold;
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
<title>申訴項目分析表</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body onkeydown="KeyDown()">
<form name="myForm" method="post">
<table border="1" style="border-right-width:0px; border-left-width:0px; border-bottom-width:0px; border-top-width:0px">
	<tr>
		<td colspan="9" align="center" class="style1" style="border-right-width:0px; border-left-width:0px; border-top-width:0px; border-top-width:0px" height="70">
		<%
	strAP="select * from apconfigure where ID=40"
	Set rsAp=conn.execute(strAP)
	If Not rsAp.eof Then
		response.write Trim(rsAp("value"))
	End If
	rsAp.close
	Set rsAp=Nothing 
	If sys_City="台南市" Or sys_City="高雄市" Or sys_City="台中市" Then
		'response.write "交通大隊"
	Else
		'response.write "交通隊"
	End If 
		%> &nbsp;<%
	'response.write Left(Trim(request("Date2")),Len(Trim(request("Date2")))-4)
	'response.write mid(Trim(request("Date2")),Len(Trim(request("Date2")))-3,2)
	response.write Trim(request("Date1")) & " ～ " & Trim(request("Date2"))
		%>&nbsp;月受理民眾交通違規申訴項目分析及員警錯誤舉發與撤銷案件統計表<%
	If Trim(request("BillType"))="1" Then
		response.write "(攔停)"
	ElseIf Trim(request("BillType"))="2" Then
		response.write "(逕舉)"
	End If 
		%>
		</td>
	</tr>
	<tr>
		<td colspan="3" align="center" class="style2">民眾申訴違規主要項目<br>(請依序排列前10項及其他)</td>
		<td colspan="3" align="center" class="style2">員警舉發錯誤主要項目<br>(請依序排列前10項及其他)</td>
		<td colspan="3" align="center" class="style2">撤銷舉發單理由<br>(請依序排列前10項及其他)</td>
	</tr>
	<tr>
		<td colspan="2" height="48" class="style2">項目</td>
		<td  class="style2">件數</td>
		<td colspan="2" class="style2">項目</td>
		<td  class="style2">件數</td>
		<td colspan="2" class="style2">項目</td>
		<td  class="style2">件數</td>
	</tr>
<%	ArgueDate1=gOutDT(Trim(request("Date1")))
	ArgueDate2=gOutDT(Trim(request("Date2")))
	strwhereD=""
	strwhereD=strwhereD&" and a.ArgueDate between "&funGetDate(ArgueDate1,0)&" and "&funGetDate(ArgueDate2,0)
	strwhereD3=strwhereD3&" and ArgueDate between "&funGetDate(ArgueDate1,0)&" and "&funGetDate(ArgueDate2,0)
	strwhereU1=" and a.RecordMemberID in (select MemberID from MemberData where UnitID in " &_
				" (select UnitID from UnitInfo where (UnitID='"&Trim(request("BillUnit"))&"') or (UnitTypeID='"&Trim(request("BillUnit"))&"'" &_
				" and ShowOrder=2)))"
	strwhereT=""
	If Trim(request("BillType"))="1" Then
		strwhereT=strwhereT&" and BillTypeID='1' "
	ElseIf Trim(request("BillType"))="2" Then
		strwhereT=strwhereT&" and BillTypeID='2' "
	End If 
	'民眾申訴違規主要項目
	strArr1a=""
	strArr1b=""
	strArr1x="0"
	str1="select a.ArguerResonID,count(*) as cnt from ArgueBase a,(select Distinct BillNo,BillTypeid from billbaseView where BillNo in (select Billno from ArgueBase where Recordstateid=0 "&strwhereD3&") "&strwhereT&") b where a.billno=b.billno and a.RecordStateID=0 and a.ArguerResonID not in (0,448) "&strwhereD&strwhereU1&" group by a.ArguerResonID order by cnt desc"
	Set rs1=conn.execute(str1)
	If Not rs1.eof Then
		For Ar1=0 To 9
			If rs1.eof Then Exit for
			If strArr1a="" then
				strArr1a=Trim(rs1("cnt"))
			Else
				strArr1a=strArr1a&"@#!"&Trim(rs1("cnt"))
			End If 
			If strArr1b="" then
				strArr1b=Trim(rs1("ArguerResonID"))
			Else
				strArr1b=strArr1b&"@#!"&Trim(rs1("ArguerResonID"))
			End If 
			If strArr1x="" then
				strArr1x=Trim(rs1("ArguerResonID"))
			Else
				strArr1x=strArr1x&","&Trim(rs1("ArguerResonID"))
			End If 
			rs1.movenext
		next
	End If 
	rs1.close
	Set rs1=Nothing

	'員警舉發錯誤主要項目
	strArr2a=""
	strArr2b=""
	strArr2x="0"
	str2="select a.ErrorID,count(*) as cnt from ArgueBase a,(select Distinct BillNo,BillTypeid from billbaseView where BillNo in (select Billno from ArgueBase where Recordstateid=0 "&strwhereD3&") "&strwhereT&") b where a.billno=b.billno and a.RecordStateID=0 and a.ErrorID not in (0,453) "&strwhereD&strwhereU1&" group by a.ErrorID order by cnt desc"
	Set rs2=conn.execute(str2)
	If Not rs2.eof then
		For Ar1=0 To 9
			If rs2.eof Then Exit for
			If strArr2a="" then
				strArr2a=Trim(rs2("cnt"))
			Else
				strArr2a=strArr2a&"@#!"&Trim(rs2("cnt"))
			End If 
			If strArr2b="" then
				strArr2b=Trim(rs2("ErrorID"))
			Else
				strArr2b=strArr2b&"@#!"&Trim(rs2("ErrorID"))
			End If 
			If strArr2x="" then
				strArr2x=Trim(rs2("ErrorID"))
			Else
				strArr2x=strArr2x&","&Trim(rs2("ErrorID"))
			End If 
			rs2.movenext
		Next 
	End If 
	rs2.close
	Set rs2=Nothing
	
	'撤銷舉發單理由
	strArr3a=""
	strArr3b=""
	strArr3x="0"
	str3="select a.DelBillReason,count(*) as cnt from ArgueBase a,(select Distinct BillNo,BillTypeid from billbaseView where BillNo in (select Billno from ArgueBase where Recordstateid=0 "&strwhereD3&") "&strwhereT&") b where a.billno=b.billno and a.RecordStateID=0 and a.DelBillReason not in (0,811) "&strwhereD&strwhereU1&" group by a.DelBillReason order by cnt desc"
	Set rs3=conn.execute(str3)
	If Not rs3.eof then
		For Ar1=0 To 9
			If rs3.eof Then Exit for
			If strArr3a="" then
				strArr3a=Trim(rs3("cnt"))
			Else
				strArr3a=strArr3a&"@#!"&Trim(rs3("cnt"))
			End If 
			If strArr3b="" then
				strArr3b=Trim(rs3("DelBillReason"))
			Else
				strArr3b=strArr3b&"@#!"&Trim(rs3("DelBillReason"))
			End If 
			If strArr3x="" then
				strArr3x=Trim(rs3("DelBillReason"))
			Else
				strArr3x=strArr3x&","&Trim(rs3("DelBillReason"))
			End If 
			rs3.movenext
		Next 
	End If 
	rs3.close
	Set rs3=Nothing

	Array1a=Split(strArr1a,"@#!")
	Array1b=Split(strArr1b,"@#!")
	Array2a=Split(strArr2a,"@#!")
	Array2b=Split(strArr2b,"@#!")
	Array3a=Split(strArr3a,"@#!")
	Array3b=Split(strArr3b,"@#!")
	
	sum1=0
	sum2=0
	sum3=0
	For Ar1=0 To 9
%>
	<tr>
		<td align="center" class="style2"><%=Ar1+1%></td>
		<td class="style2" ><%
		If Ar1 <= UBound(Array1b) Then
			if Not IsNull(Array1b(Ar1)) And Trim(Array1b(Ar1))<>"" Then
				strA1="select * from code where id="&Trim(Array1b(Ar1))
				Set rsA1=conn.execute(strA1)
				If Not rsA1.eof Then
					response.write Trim(rsA1("Content"))
				End If
				rsA1.close
				Set rsA1=Nothing 
			End if
		Else
			response.write "&nbsp;"
		End If 
		
		%></td>
		<td align="center" class="style2"><%
		If Ar1 <= UBound(Array1a) Then
			response.write Array1a(Ar1)
			sum1=sum1+CDbl(Array1a(Ar1))
		Else
			response.write "&nbsp;"
		End If 
		%></td>
		<td align="center" class="style2" <%
		If Ar1 <= UBound(Array1b) Then
			If Trim(Array1b(Ar1))="601" Then
				response.write " height=""90"""
			Else
				response.write " height=""48"""
			End If 
		Else
			response.write " height=""48"""
		End If 
		%>><%=Ar1+1%></td>
		<td class="style2" ><%
		If Ar1 <= UBound(Array2b) Then
			if Not IsNull(Array2b(Ar1)) And Trim(Array2b(Ar1))<>"" Then
				strA2="select * from code where id="&Trim(Array2b(Ar1))
				Set rsA2=conn.execute(strA2)
				If Not rsA2.eof Then
					response.write Trim(rsA2("Content"))
				End If
				rsA2.close
				Set rsA2=Nothing 
			End if
		Else
			response.write "&nbsp;"
		End If 
		
		%></td>
		<td align="center" class="style2" ><%
		If Ar1 <= UBound(Array2a) Then
			response.write Array2a(Ar1)
			sum2=sum2+CDbl(Array2a(Ar1))
		Else
			response.write "&nbsp;"
		End If 
		%></td>
		<td  align="center" class="style2"><%=Ar1+1%></td>
		<td class="style2" ><%
		If Ar1 <= UBound(Array3b) Then
			if Not IsNull(Array3b(Ar1)) And Trim(Array3b(Ar1))<>"" Then
				strA3="select * from code where id="&Trim(Array3b(Ar1))
				Set rsA3=conn.execute(strA3)
				If Not rsA3.eof Then
					response.write Trim(rsA3("Content"))
				End If
				rsA3.close
				Set rsA3=Nothing 
			End if
		Else
			response.write "&nbsp;"
		End If 
		
		%></td>
		<td align="center" class="style2"><%
		If Ar1 <= UBound(Array3a) Then
			response.write Array3a(Ar1)
			sum3=sum3+CDbl(Array3a(Ar1))
		Else
			response.write "&nbsp;"
		End If 
		%></td>
	</tr>
<%
	Next
%>
	<tr>
		<td width="25" align="center" height="48" class="style2"><%="11"%></td>
		<td width="140" class="style2" ><%
			response.write "其他"
		%></td>
		<td width="60" align="center" class="style2"><%
		str1="select count(*) as cnt from ArgueBase a,(select Distinct BillNo,BillTypeid from billbaseView where BillNo in (select Billno from ArgueBase where Recordstateid=0 "&strwhereD3&") "&strwhereT&") b where a.billno=b.billno and a.RecordStateID=0 and a.ArguerResonID not in ("&strArr1x&") "&strwhereD&strwhereU1
		Set rs1=conn.execute(str1)
		If Not rs1.eof Then
			response.write Trim(rs1("cnt"))
			sum1=sum1+CDbl(Trim(rs1("cnt")))
		Else
			response.write "&nbsp;"
		End If 
		rs1.close
		Set rs1=Nothing
		%></td>
		<td width="25" align="center" class="style2" ><%="11"%></td>
		<td width="140" class="style2"><%
			response.write "其他"		
		%></td>
		<td width="60" align="center" class="style2" ><%
		str1="select count(*) as cnt from ArgueBase a,(select Distinct BillNo,BillTypeid from billbaseView where BillNo in (select Billno from ArgueBase where Recordstateid=0 "&strwhereD3&") "&strwhereT&") b where a.billno=b.billno and a.RecordStateID=0 and a.ErrorID not in ("&strArr2x&") "&strwhereD&strwhereU1
		Set rs1=conn.execute(str1)
		If Not rs1.eof Then
			response.write Trim(rs1("cnt"))
			sum2=sum2+CDbl(Trim(rs1("cnt")))
		Else
			response.write "&nbsp;"
		End If 
		rs1.close
		Set rs1=Nothing
		%></td>
		<td width="25" align="center" class="style2"><%="11"%></td>
		<td width="140" class="style2" ><%
			response.write "其他"
		%></td>
		<td width="60" align="center" class="style2"><%
		str1="select count(*) as cnt from ArgueBase a,(select Distinct BillNo,BillTypeid from billbaseView where BillNo in (select Billno from ArgueBase where Recordstateid=0 "&strwhereD3&") "&strwhereT&") b where a.billno=b.billno and a.RecordStateID=0 and a.DelBillReason not in ("&strArr3x&") "&strwhereD&strwhereU1
		Set rs1=conn.execute(str1)
		If Not rs1.eof Then
			response.write Trim(rs1("cnt"))
			sum3=sum3+CDbl(Trim(rs1("cnt")))
		Else
			response.write "&nbsp;"
		End If 
		rs1.close
		Set rs1=Nothing
		%></td>
	</tr>
	<tr>
		<td colspan="2" height="48" class="style2">合計</td>
		<td align="center" class="style2"><%=sum1%></td>
		<td colspan="2" class="style2">合計</td>
		<td align="center" class="style2"><%=sum2%></td>
		<td colspan="2" class="style2">合計</td>
		<td align="center" class="style2"><%=sum3%></td>
	</tr>
	<tr>
		<td colspan="9" class="style3" style="border-right-width:0px;border-left-width:0px;border-bottom-width:0px">
			備註： &nbsp;<strong><%
	response.write Left(Trim(request("Date2")),Len(Trim(request("Date2")))-4)
		%></strong>&nbsp;年 &nbsp;至 &nbsp;<strong><%
	response.write mid(Trim(request("Date2")),Len(Trim(request("Date2")))-3,2)
		%></strong>&nbsp;月受理民眾申訴案件中，經查證為員警舉發錯誤，確有疏失者共 &nbsp;<strong><%=sum2&"&nbsp;"%></strong>件，占該期申訴總數 &nbsp;<strong><%
			If sum1<>0 Then
				response.write Round(sum2/sum1,3) &"&nbsp;"
			Else
				response.write "0"&"&nbsp;"
			End If 
			
			%></strong>。與 &nbsp;<strong><%
			ArgueDate1b=DateAdd("yyyy",-1,gOutDT(Trim(request("Date1"))))
			ArgueDate2b=DateAdd("yyyy",-1,gOutDT(Trim(request("Date2"))))
			response.write Year(ArgueDate2b)-1911
			%></strong>&nbsp;年同期(受理申訴 &nbsp;<strong><%			
			strwhereD2=""
			Osum1=0
			strwhereD2=strwhereD2&" and a.ArgueDate between "&funGetDate(ArgueDate1b,0)&" and "&funGetDate(ArgueDate2b,0)
			strwhereD4=strwhereD4&" and ArgueDate between "&funGetDate(ArgueDate1b,0)&" and "&funGetDate(ArgueDate2b,0)
			strO1="select count(*) as cnt from ArgueBase a,(select Distinct BillNo,BillTypeid from billbaseView where BillNo in (select Billno from ArgueBase where Recordstateid=0 "&strwhereD4&") "&strwhereT&") b where a.billno=b.billno and a.RecordStateID=0 and a.ArguerResonID<>0 "&strwhereD2&strwhereU1
			Set rsO1=conn.execute(strO1)
			If Not rsO1.eof Then
				response.write Trim(rsO1("cnt"))
				Osum1=CDbl(Trim(rsO1("cnt")))
			End If
			rsO1.close
			Set rsO1=Nothing
			
			%></strong>&nbsp;件、舉發錯誤 &nbsp;<strong><%
			Osum2=0
			strO1="select count(*) as cnt from ArgueBase a,(select Distinct BillNo,BillTypeid from billbaseView where BillNo in (select Billno from ArgueBase where Recordstateid=0 "&strwhereD4&") "&strwhereT&") b where a.billno=b.billno and a.RecordStateID=0 and a.ErrorID<>0 "&strwhereD2&strwhereT&strwhereU1
			Set rsO1=conn.execute(strO1)
			If Not rsO1.eof Then
				response.write Trim(rsO1("cnt"))
				Osum2=CDbl(Trim(rsO1("cnt")))
			End If
			rsO1.close
			Set rsO1=Nothing
			
			%></strong>&nbsp;件)比較，受理民眾申訴案件<strong><%
			If Osum1 > sum1 Then
				response.write " 減少 "
				response.write Osum1-sum1
			Else
				response.write " 增加 "
				response.write sum1-Osum1
			End If 
			%></strong>&nbsp;件，員警舉發錯誤，確有疏失者<strong><%
			If Osum2 > sum2 Then
				response.write " 減少 "
				response.write Osum2-sum2
			Else
				response.write " 增加 "
				response.write sum2-Osum2
			End If 
			%></strong> &nbsp;件。

			<br><br>填單人： &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;單位主管：
		</td>
	</tr>
</table>

</form>
</body>
</html>
<%conn.close%>