<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<%
'fMnoth=month(now)
'if fMnoth<10 then fMnoth="0"&fMnoth
'fDay=day(now)
'if fDay<10 then	fDay="0"&fDay
'fname=year(now)&fMnoth&fDay&"_移送清冊.xls"
'Response.AddHeader "Content-Disposition", "filename="&fname
'response.contenttype="application/x-msexcel; charset=MS950" 

%>
<html>
<head>

<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style3 {font-family:新細明體; color=0044ff; line-height:19px; font-size: 15px}
.pageprint {
  margin-left: 7mm;
  margin-right: 5.08mm;
  margin-top: 5.08mm;
  margin-bottom: 5.08mm;
}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>退件清冊</title>

<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<%
Server.ScriptTimeout = 18000
Response.flush

%>
<%

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
If Not rsCity.eof then
	sys_City=trim(rsCity("value"))
End If 
rsCity.close
set rsCity=Nothing

strCity="select value from Apconfigure where id=49"
set rsCity=conn.execute(strCity)
If Not rsCity.eof then
	sys_Unit1=trim(rsCity("value"))
End If 
rsCity.close
set rsCity=Nothing

UserUnitTypeID=""
strU="select UnitTypeID from UnitInfo where UnitID='"&Trim(session("Unit_ID"))&"'"
Set rsU=conn.execute(strU)
If Not rsU.eof Then
	UserUnitTypeID=Trim(rsU("UnitTypeID"))
End If 
rsU.close
Set rsU=Nothing 






	strwhere="select * from billbase where recordstateid=-1 and exists (" & _
		"select Billsn from Dcilog "&Trim(request("strDCISQL"))&" and Billsn=billbase.Sn " & _
		"" & _
		")"



	strSQL=strwhere
	'response.write strSQL	
	set rs=conn.execute(strSQL)


cnt=0
%>
<table width="98%" border="1" cellpadding="1" cellspacing="0">
	<tr>
		<td height="33" colspan='11'>退件清冊

		</td>
	</tr>
	<tr align="center" class="font10">
		<td ></td>
		<td >車號</td>
		<!-- <td >車種</td> -->
		<td >違規日</td>
		<td >時間</td>	
		<td >違規地點</td>
		<td >法條</td>
		<td >舉發單位</td>
		<td >舉發員警</td>
		<td >退件原因</td>
		<td >退件日期</td>
		<!--<td >檢舉日</td>
		<td >專案代碼</td>-->
	</tr><%
		while Not rs.eof

			response.write "<tr align='center' "
			response.write ">"
%>
		<td><%
		cnt=cnt+1
		response.write cnt

		BillStatusTemp=""
		BillNoTemp=""
		chname=""
		chRule=""

		BillNoTemp=Trim(rs("BillNo"))
'		if rs("BillMem1")<>"" then	chname=rs("BillMem1")
'		if rs("BillMem2")<>"" then	chname=chname&"/"&rs("BillMem2")
'		if rs("BillMem3")<>"" then	chname=chname&"/"&rs("BillMem3")
'		if rs("BillMem4")<>"" then	chname=chname&"/"&rs("BillMem4")
		if rs("Rule1")<>"" then chRule=rs("Rule1")
		if rs("Rule2")<>"" then chRule=chRule&"<br>"&rs("Rule2")
		if rs("Rule3")<>"" then chRule=chRule&"<br>"&rs("Rule3")
		%></td>

		<td >
<%
		If Trim(rs("CarNo"))<>"" Then
			response.write Trim(rs("CarNo"))&"&nbsp;"
		Else
			response.write "&nbsp;"
		End if
		%>
		</td>
		<!-- <td><%
			if trim(rs("CarSimpleID"))="1" then
				response.write "<span class=""style5"">汽車</span>"
			elseif trim(rs("CarSimpleID"))="2" then
				response.write "<span class=""style5"">拖車</span>"
			elseif trim(rs("CarSimpleID"))="3" then
				response.write "<span class=""style5"">重機</span>"
			elseif trim(rs("CarSimpleID"))="4" then
				response.write "<span class=""style5"">輕機</span>"
			elseif trim(rs("CarSimpleID"))="5" then
				response.write "<span class=""style5"">動力機械</span>"
			elseif trim(rs("CarSimpleID"))="6" then
				response.write "<span class=""style5"">臨時車牌</span>"
			elseif trim(rs("CarSimpleID"))="7" then
				response.write "<span class=""style5"">試車牌</span>"
			end if
		%></td> -->
		<td><%
		If trim(rs("IllegalDate"))<>"" then
			response.write gInitDT(trim(rs("IllegalDate")))
		End If 
		%></td>
		<td><%
		If trim(rs("IllegalDate"))<>"" then
			response.write right("00"&hour(rs("IllegalDate")),2)&right("00"&minute(rs("IllegalDate")),2)
		End If 
		%></td>
		<td align="left"><%
		If Trim(rs("IllegalAddress"))<>"" Then
			response.write Trim(rs("IllegalAddress"))
		Else
			response.write "&nbsp;"
		End if
		%></td>
		<td><%
		response.write chRule
		%></td>
		<td><%
		If Trim(rs("BillUnitID"))<>"" Then
			strBU="select UnitName from UnitInfo where UnitID='"&Trim(rs("BillUnitID"))&"'"
			Set rsBU=conn.execute(strBU)
			If Not rsBU.eof Then
				response.write Trim(rsBU("UnitName"))
			End If
			rsBU.close
			Set rsBU=Nothing 
		End if
		%></td>
		<td><%
		If Trim(rs("BillMemID1"))<>"" Then
			strM="select loginid,chname from memberdata where memberid="&Trim(rs("BillMemID1"))
			Set rsM=conn.execute(strM)
			If Not rsM.eof Then
				response.write Trim(rsM("chname"))
			End If 
			rsM.close
			Set rsM=Nothing 
		End If 
		If Trim(rs("BillMemID2"))<>"" Then
			strM="select loginid,chname from memberdata where memberid="&Trim(rs("BillMemID2"))
			Set rsM=conn.execute(strM)
			If Not rsM.eof Then
				response.write "<br>" & Trim(rsM("chname"))
			End If 
			rsM.close
			Set rsM=Nothing 
		End If 
		If Trim(rs("BillMemID3"))<>"" Then
			strM="select loginid,chname from memberdata where memberid="&Trim(rs("BillMemID3"))
			Set rsM=conn.execute(strM)
			If Not rsM.eof Then
				response.write "<br>" & Trim(rsM("chname"))
			End If 
			rsM.close
			Set rsM=Nothing 
		End If 
		If Trim(rs("BillMemID4"))<>"" Then
			strM="select loginid,chname from memberdata where memberid="&Trim(rs("BillMemID4"))
			Set rsM=conn.execute(strM)
			If Not rsM.eof Then
				response.write "<br>" & Trim(rsM("chname"))
			End If 
			rsM.close
			Set rsM=Nothing 
		End If 

		response.write chname
		%></td>
		<td><%
		ReturnReason=""
		ReturnDate=""
		strR="select * from BillDeleteReason where billsn="& Trim(rs("Sn"))
		Set rsR=conn.execute(strR)
		If Not rsR.eof Then
			ReturnReason=Trim(rsR("Note"))
			ReturnDate=gInitDT(trim(rsR("DelDate")))
		End If 
		rsR.close
		Set rsR=Nothing 
		If ReturnReason<>"" then
			response.write ReturnReason
		Else
			response.write "&nbsp;"
		End If 
		%></td>

		<td><%
		If ReturnDate<>"" then
			response.write ReturnDate&"&nbsp;"
		Else
			response.write "&nbsp;"
		End If 

		
		%></td>

<%
			response.write ""
			response.write "</tr>"
			rs.movenext
		wend
		rs.close
		set rs=nothing
		%>
</table>


</form>
</body>
</html>
<script language="javascript">
function DP(){
	window.focus();
	window.print();
}

window.print();

</script>
<%conn.close%>