<!-- #include file="Common\db.ini" -->
<!-- #include file="Common\AllFunction.inc" -->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<head>
<meta http-equiv="Content-Language" content="zh-tw">

<title>
</title>
</head>

<table border="1" width="100%" id="table1">
	<tr bgcolor="#FFCC33">
		<td>指令</td>
		<td>資源佔用量</td>
		<td>使用者</td>
		<td>程式</td>
	</tr>
<%
	strsql="select se.command,ss.value CPU,se.username,se.program from v$sesstat ss,v$session se " &_
		" where ss.statistic# in (select statistic# from v$statname where name='CPU used by this session') " &_
		" and se.sid=ss.sid and ss.sid>6 " &_
		" and (UserName like '%TJDB%' or UserName like '%DEDB%')"
	set rs1=conn.execute(strsql)
	If Not rs1.Bof Then rs1.MoveFirst 
	While Not rs1.Eof
%>
	<tr>
		<td><%=trim(rs1("command"))%></td>
		<td><%=trim(rs1("CPU"))%></td>
		<td><%="交通事故系統"%></td>
		<td><%=trim(rs1("program"))%></td>
	</tr>
<%
		rs1.MoveNext
		Wend
	rs1.close
	set rs1=nothing
	conn.close
	set conn=nothing
%>
</table>
