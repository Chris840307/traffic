<!-- #include file="Common\db.ini" -->
<!-- #include file="Common\AllFunction.inc" -->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<head>
<meta http-equiv="Content-Language" content="zh-tw">

<title>各單位舉發單職名章主管人員
</title>
</head>

<table border="1" width="550" id="table1">
	<tr><td bgcolor="#FFCC66" colspan="2">各單位舉發單職名章主管人員</td></tr>
	<tr bgcolor="#FFCC33">
		<td width="30%">職稱</td>
		<td width="30%">姓名</td>
	</tr>
<%
	strsql="select distinct c.unitname,c.unitid from (select UnitID,ChName,JobID" &_
		" from MemberData where AccountStateID=0 and RecordStateID=0 " &_
		" and JobID in(303,304,305,307,314,318,1936,1937,1935,1938)) a," &_
		" (select ID,showorder,Content from Code where TypeID=4 ) b, " &_
		" (select Unitid,UnitName from Unitinfo) c " &_
		" where a.JobID=b.ID and a.unitid=c.unitid order by unitid"
	set rs1=conn.execute(strsql)
	If Not rs1.Bof Then rs1.MoveFirst 
	While Not rs1.Eof
%>
	<tr>
		<td colspan="2"  bgcolor="#FFFF99"><%=trim(rs1("unitname"))%></td>
	</tr>
<%
			strsql2="select a.unitid,c.unitname,a.ChName,b.Content,b.ID,b.showorder from " &_
				" (select UnitID,ChName,JobID" &_
				" from MemberData where AccountStateID=0 and RecordStateID=0 " &_
				" and JobID in(303,304,305,307,314,318,1936,1937,1935,1938,1838)) a," &_
				" (select ID,showorder,Content from Code where TypeID=4 ) b, " &_
				" (select Unitid,UnitName from Unitinfo) c " &_
				" where a.JobID=b.ID and a.unitid=c.unitid" &_
				" and a.unitid='"&Trim(rs1("unitid"))&"'" &_
				" order by unitid,showorder,b.id"
			set rs2=conn.execute(strsql2)
			If Not rs2.eof Then 
		%>
			<tr>
				<td width="30%"><%=trim(rs2("Content"))%></td>
				<td width="30%"><%=trim(rs2("ChName"))%></td>
			</tr>
		<%
			End if
			rs2.close
			set rs2=nothing

		rs1.MoveNext
		Wend
	rs1.close
	set rs1=nothing
	conn.close
	set conn=nothing
%>
</table>
