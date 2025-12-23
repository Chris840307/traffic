<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="Traffic/Common/AllFunction.inc"-->
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
'sys_City="雲林縣"
%>

<%if sys_City="台中縣" Or sys_City="台中市" Or sys_City="南投縣" Or sys_City="基隆市" Or sys_City="澎湖縣" then %>
<!--#include virtual="Traffic/Common/OlddbAccess.ini"-->
<%else%>
<!--#include virtual="Traffic/Common/OldbAccessHualien.ini"-->
<%end if%>
<%
Server.ScriptTimeout = 8648000
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_違規道路障礙罰單統計清冊.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 

	sql = "select sysdate from Dual"
	Set RSSystem = Conn.Execute(sql)
	DBDate = RSSystem("sysdate")

	sql = "select UnitName from UnitInfo where UnitID= '" & Session("Unit_ID") & "'"
	Set RSSystem = Conn.Execute(sql)
	if Not RSSystem.Eof Then
		printUnit = RSSystem("UnitName")
	End If	

	If sys_City="雲林縣" Or sys_City="花蓮縣" Then 
		If InStr(request("strwhere"),"SEQNO")=0 Then 
				If Trim(request("CloseFlag"))<>"" Then 
					strSQL="select FSEQ,RBDATE,IDATE,ITIME,INAME,IIDNO,RULEF1,RULEF2,RULEF3,RULEF4,ARVDATE from FMaster where 1<>1"
					set rs=conn1.execute(strSQL)
					strSQL="select FSEQ,RBDATE,IDATE,ITIME,INAME,IIDNO,RULEF1,RULEF2,RULEF3,RULEF4,ARVDATE from FMaster where 1<>1"
					set rs2=conn2.execute(strSQL)
					strSQL="select FSEQ,RBDATE,IDATE,ITIME,INAME,IIDNO,RULEF1,RULEF2,RULEF3,RULEF4,ARVDATE from FMaster where 1<>1"
					set rs3=conn3.execute(strSQL)
				Else
					strSQL="select FSEQ,RBDATE,IDATE,ITIME,INAME,IIDNO,RULEF1,RULEF2,RULEF3,RULEF4,ARVDATE from FMaster " & request("strwhere")
					set rs=conn1.execute(strSQL)
					strSQL="select FSEQ,RBDATE,IDATE,ITIME,INAME,IIDNO,RULEF1,RULEF2,RULEF3,RULEF4,ARVDATE from FMaster " & request("strwhere")
					set rs2=conn2.execute(strSQL)
					strSQL="select FSEQ,RBDATE,IDATE,ITIME,INAME,IIDNO,RULEF1,RULEF2,RULEF3,RULEF4,ARVDATE from FMaster " & request("strwhere")
					set rs3=conn3.execute(strSQL)
				End If
		End if
	Else
		strSQL="select FSEQ,RBDATE,IDATE,ITIME,INAME,IIDNO,RULEF1,RULEF2,RULEF3,RULEF4,ARVDATE from FMaster " & request("strwhere")
		set rs=conn1.execute(strSQL)
		strSQL="select FSEQ,RBDATE,IDATE,ITIME,INAME,IIDNO,RULEF1,RULEF2,RULEF3,RULEF4,ARVDATE from FMaster " & request("strwhere")
		set rs2=conn2.execute(strSQL)
		strSQL="select FSEQ,RBDATE,IDATE,ITIME,INAME,IIDNO,RULEF1,RULEF2,RULEF3,RULEF4,ARVDATE from FMaster " & request("strwhere")
		set rs3=conn3.execute(strSQL)
	End if

	If sys_City="雲林縣" Or sys_City="花蓮縣" Then 
		If Trim(request("CloseFlag"))<>"" Then 
			strSQL="select FSEQ,SEQNO,RBDATE,IDATE,ITIME,INAME,IIDNO,RULEF1,RULEF2,RULEF3,RULEF4,ARVDATE from FMaster_s " & request("strwhere") & " and CloseFlag='"&request("CloseFlag")&"'"
		Else
			strSQL="select FSEQ,SEQNO,RBDATE,IDATE,ITIME,INAME,IIDNO,RULEF1,RULEF2,RULEF3,RULEF4,ARVDATE from FMaster_s " & request("strwhere")
		End if
	Else
		strSQL="select FSEQ,SEQNO,RBDATE,IDATE,ITIME,INAME,IIDNO,RULEF1,RULEF2,RULEF3,RULEF4,ARVDATE from FMaster_s " & request("strwhere")
	End if
	set rs4=conn1.execute(strSQL)


%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>違規道路障礙罰單統計清冊</title>
</head>

<body>

<p align="center"><font size="6" face="標楷體"><%=Sys_City%>警察局<%=Replace(printUnit,Sys_City,"")%>違規道路障礙罰單統計清冊</font></p>
<br><font face="標楷體">製表單位：<%=Replace(printUnit,Sys_City,"")%></font>
<br><font face="標楷體">製表人員：<%=Session("Ch_Name")%></font>
<br><font face="標楷體">製表時間：<%=gInitDT(DBDate)%></font>
<table border="1" width="100%" id="table1">
	<tr>
		<td align="center"><font face="標楷體">編號</font></td>
		<td align="center"><font face="標楷體">告發單號</font></td>
		<td align="center"><font face="標楷體">告發單日期</font></td>
		<td align="center"><font face="標楷體">違規日期</font></td>
		<td align="center"><font face="標楷體">時間</font></td>
		<td align="center"><font face="標楷體">姓名</font></td>
		<td align="center"><font face="標楷體">違規人證號</font></td>
		<td align="center"><font face="標楷體">法條代碼</font></td>
		<td align="center"><font face="標楷體">應到案日期</font></td>
		<td align="center"><font face="標楷體">慢車行人序號</font></td>
		<td colspan="2" align="center"><font face="標楷體">罰款金額（最高罰）</font></td>
	</tr>


	<%
				i=0
			If InStr(request("strwhere"),"SEQNO")=0 Then 
					while Not rs.eof
					 i=i+1  
						response.write "<td><font face=""標楷體"">&nbsp;"&i&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs("FSEQ")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs("RBDATE")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs("IDATE")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs("ITIME")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs("INAME")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs("IIDNO")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs("RULEF1")&" "&rs("RULEF2")&" "&rs("RULEF3")&" "&rs("RULEF4")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs("ARVDATE")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;</font></td>"
						response.write "<td><font face=""標楷體"">罰鍰新台幣</font></td>"
						response.write "<td><font face=""標楷體"">陸百元整</font></td>"

	 					response.write "</tr>"
						rs.movenext
					Wend
					while Not rs2.eof
					 i=i+1  
						response.write "<td><font face=""標楷體"">&nbsp;"&i&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs2("FSEQ")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs2("RBDATE")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs2("IDATE")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs2("ITIME")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs2("INAME")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs2("IIDNO")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs2("RULEF1")&" "&rs2("RULEF2")&" "&rs2("RULEF3")&" "&rs2("RULEF4")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs2("ARVDATE")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;</font></td>"
						response.write "<td><font face=""標楷體"">罰鍰新台幣</font></td>"
						response.write "<td><font face=""標楷體"">陸百元整</font></td>"

	 					response.write "</tr>"
						rs2.movenext
					wend					
					while Not rs3.eof
					 i=i+1  
						response.write "<td><font face=""標楷體"">&nbsp;"&i&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs3("FSEQ")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs3("RBDATE")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs3("IDATE")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs3("ITIME")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs3("INAME")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs3("IIDNO")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs3("RULEF1")&" "&rs3("RULEF2")&" "&rs3("RULEF3")&" "&rs3("RULEF4")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs3("ARVDATE")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;</font></td>"
						response.write "<td><font face=""標楷體"">罰鍰新台幣</font></td>"
						response.write "<td><font face=""標楷體"">陸百元整</font></td>"

	 					response.write "</tr>"
						rs3.movenext
					Wend
				End if 
					while Not rs4.eof
					 i=i+1  
						response.write "<td><font face=""標楷體"">&nbsp;"&i&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs4("FSEQ")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs4("RBDATE")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs4("IDATE")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs4("ITIME")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs4("INAME")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs4("IIDNO")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs4("RULEF1")&" "&rs4("RULEF2")&" "&rs4("RULEF3")&" "&rs4("RULEF4")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs4("ARVDATE")&"</font></td>"
						response.write "<td><font face=""標楷體"">&nbsp;"&rs4("SEQNO")&"</font></td>"
						response.write "<td><font face=""標楷體"">罰鍰新台幣</font></td>"
						response.write "<td><font face=""標楷體"">陸百元整</font></td>"

	 					response.write "</tr>"
						rs4.movenext
					wend					
					
					Set rs1=Nothing
					Set rs2=nothing					
					Set rs3=Nothing
					Set rs4=nothing					
					Set conn1=Nothing
					Set conn2=nothing					
					Set conn3=nothing
					%>
	</tr>
</table>

</body>
<script>
window.close();
</script>
</html>
