<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title>舉發單查詢</title>
<%
'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post" onsubmit="return funBillQry();">  
儲存成功!<br>
<%If sys_City="台南市" then%>
下列舉發案件為三個月內，在<%=Trim(request("IllegalZipName"))%>，本車號民眾檢舉案件，<br>
<%else%>
下列舉發案件為一周內，本車號民眾檢舉案件，<br>
<%End if%>
請確認是否重複舉發!<br>
(如確認本案件要舉發，不須理會本訊息)
		<table width='100%' border='1' align="left" cellpadding="1">
<%
If Trim(request("BillBaseTmpFlag"))="1" Then
	strQry="select SN,BillNo,CarNo,IllegalDate,IllegalAddress from BillBaseTmp where Sn="&Trim(request("BillSN"))
else
	strQry="select SN,BillNo,CarNo,IllegalDate,IllegalAddress from BillBase where Sn="&Trim(request("BillSN"))
End If 
	set rsQry=conn.execute(strQry)
	if not rsQry.eof then
		CarNoTemp=Trim(rsQry("CarNo"))
		IllegalAddressTemp=Trim(rsQry("IllegalAddress"))
		illegalDateTmp=year(rsQry("IllegalDate"))&"/"&month(rsQry("IllegalDate"))&"/"&Day(rsQry("IllegalDate"))&" "&hour(rsQry("IllegalDate"))&":"&minute(rsQry("IllegalDate"))&":00"
		
		illegalDate1=DateAdd("d",-7,illegalDateTmp)
		illegalDate2=DateAdd("d",7,illegalDateTmp)
		
	end if
	rsQry.close
	set rsQry=Nothing

	strIllDate=" and IllegalDate between TO_DATE('"&year(illegalDate1)&"/"&month(illegalDate1)&"/"&day(illegalDate1)&" "&Hour(illegalDate1)&":"&minute(illegalDate1)&":00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&year(illegalDate2)&"/"&month(illegalDate2)&"/"&day(illegalDate2)&" "&Hour(illegalDate2)&":"&minute(illegalDate2)&":59','YYYY/MM/DD/HH24/MI/SS')"
	
	strIllDate = strIllDate &  " and JurgeDay is not null "

	If sys_City="台南市" And Trim(request("IllegalZipName"))<>"" Then
		strIllDate=" and IllegalAddress like '"&Trim(request("IllegalZipName"))&"%' and (Rule1 like '55%' or Rule1 like '56%' or Rule2 like '55%' or Rule2 like '56%')"
	End If 

	strChk="select BillNo,CarNo,Rule1,Rule2,IllegalAddress,(select UnitName from UnitInfo where UnitID=BillUnitID) as UnitName,Rule1,IllegalDate,JurgeDay" &_
		" from Billbase where sn<>"&Trim(request("BillSN")) &_
		" and carno='"&UCase(CarNoTemp)&"'" &_
		" and Recordstateid=0 " & strIllDate & " order by IllegalDate"
	set rs1=conn.execute(strChk)
	If Not rs1.Bof Then rs1.MoveFirst 
	While Not rs1.Eof
%>
		<tr><td>
		單號:<%
		If Trim(rs1("BillNo"))<>"" Then 
			response.write Trim(rs1("BillNo"))
		Else
			response.write "未入案"
		End If 
		%>,車號:<%=Trim(rs1("CarNo"))%><br>
		違規時間:<%=Year(rs1("IllegalDate"))-1911&"/"&month(rs1("IllegalDate"))&"/"&day(rs1("IllegalDate"))&" "&hour(rs1("IllegalDate"))&":"&minute(rs1("IllegalDate"))%>&nbsp; &nbsp; 
		檢舉日期:<%=Year(rs1("JurgeDay"))-1911&"/"&month(rs1("JurgeDay"))&"/"&day(rs1("JurgeDay"))%><br>
		違規地點:<%=Trim(rs1("IllegalAddress"))%><br>
		違規法條:<%
		response.write Trim(rs1("Rule1"))
		strR1="select * from Law where itemid='"&Trim(rs1("Rule1"))&"' and version=2"
		Set rsR1=conn.execute(strR1)
		If Not rsR1.eof Then
			response.write " " & Trim(rsR1("IllegalRule"))
		End If 
		rsR1.close
		Set rsR1=Nothing 
		%><%
		If Trim(rs1("Rule2"))<>"" Then
			response.write "<br>"&Trim(rs1("Rule2"))
			strR2="select * from Law where itemid='"&Trim(rs1("Rule2"))&"' and version=2"
			Set rsR2=conn.execute(strR2)
			If Not rsR2.eof Then
				response.write " " & Trim(rsR2("IllegalRule"))
			End If 
			rsR2.close
			Set rsR2=Nothing 
		End If 
		%><br>
		舉發單位:<%=Trim(rs1("UnitName"))%>
		</td></tr>
<%
	rs1.MoveNext
	Wend
	rs1.close
	set rs1=nothing
%>
		</table>		
	</form>
<%
conn.close
set conn=nothing
%>
</body>
<script language="JavaScript">

</script>
</html>
