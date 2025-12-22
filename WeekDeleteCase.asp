<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
#Layer1 {
	position:absolute;
	width:209px;
	height:38px;
	z-index:2;
	top: 13px;
}
.style1 {font-size: 16px}
.style2 {font-size: 13px; }
#Layer2 {
	position:absolute;
	width:566px;
	height:38px;
	z-index:3;
	top: 15px;
}
#LayerTime {
	position:absolute;
	width:311px;
	height:34px;
	z-index:1;
}
#Layer151 {
	position:absolute;
	width:209px;
	height:38px;
	z-index:2;
}
-->
</style>
<head>
<!--#include virtual="traffic/Common/css.txt"-->
<title>七日內刪除案件列表</title>
</head>
<body leftmargin="5" topmargin="5" marginwidth="0" marginheight="0">
<table width='720' border='1' align="center" >
	<tr bgcolor="#FFCC33">
		<td><strong>七日內刪除案件列表</strong></td>
	</tr>
	<tr bgcolor="#FFFFCC">
		<td><%'1
		response.write(year(now)-1911&"/"&month(now)&"/"&day(now))
		%></td>
	</tr>
	<tr>
		<td><span class="style1">
<%	CaseSn=0
	strDel1="select * from Log where ActionDate between " &_
		"TO_DATE('"&year(now)&"/"&month(now)&"/"&day(now)&" 0:0:0','YYYY/MM/DD/HH24/MI/SS')" &_
		" and TO_DATE('"&year(now)&"/"&month(now)&"/"&day(now)&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')" &_
		" and Typeid=352 order by ActionMemberID"
	set rsDel1=conn.execute(strDel1)
	If Not rsDel1.Bof Then rsDel1.MoveFirst 
	While Not rsDel1.Eof
		if inStr(trim(rsDel1("ActionContent")),"單號:G")>0 and right(trim(rsDel1("ActionContent")),2)<>",," then
			CaseSn=CaseSn+1
			response.write CaseSn&". 刪除人:"&trim(rsDel1("ActionChName"))&"  "&replace(trim(rsDel1("ActionContent")),"舉發單刪除","")&"<br>"
		end if
		rsDel1.MoveNext
	Wend
	rsDel1.close
	set rsDel1=nothing
%>
		</span></td>
	</tr>

	<tr bgcolor="#FFFFCC">
		<td><%'2
		response.write(year(DateAdd("D",-1,now))-1911&"/"&month(DateAdd("D",-1,now))&"/"&day(DateAdd("D",-1,now)))
		%></td>
	</tr>
	<tr>
		<td><span class="style1">
<%	CaseSn=0
	strDel1="select * from Log where ActionDate between " &_
		"TO_DATE('"&year(DateAdd("D",-1,now))&"/"&month(DateAdd("D",-1,now))&"/"&day(DateAdd("D",-1,now))&" 0:0:0','YYYY/MM/DD/HH24/MI/SS')" &_
		" and TO_DATE('"&year(DateAdd("D",-1,now))&"/"&month(DateAdd("D",-1,now))&"/"&day(DateAdd("D",-1,now))&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')" &_
		" and Typeid=352 order by ActionMemberID"
	set rsDel1=conn.execute(strDel1)
	If Not rsDel1.Bof Then rsDel1.MoveFirst 
	While Not rsDel1.Eof
		if inStr(trim(rsDel1("ActionContent")),"單號:G")>0 and right(trim(rsDel1("ActionContent")),2)<>",," then
			CaseSn=CaseSn+1
			response.write CaseSn&". 刪除人:"&trim(rsDel1("ActionChName"))&"  "&replace(trim(rsDel1("ActionContent")),"舉發單刪除","")&"<br>"
		end if
		rsDel1.MoveNext
	Wend
	rsDel1.close
	set rsDel1=nothing
%>
		</span></td>
	</tr>

	<tr bgcolor="#FFFFCC">
		<td><%'3
		response.write(year(DateAdd("D",-2,now))-1911&"/"&month(DateAdd("D",-2,now))&"/"&day(DateAdd("D",-2,now)))
		%></td>
	</tr>
	<tr>
		<td><span class="style1">
<%	CaseSn=0
	strDel1="select * from Log where ActionDate between " &_
		"TO_DATE('"&year(DateAdd("D",-2,now))&"/"&month(DateAdd("D",-2,now))&"/"&day(DateAdd("D",-2,now))&" 0:0:0','YYYY/MM/DD/HH24/MI/SS')" &_
		" and TO_DATE('"&year(DateAdd("D",-2,now))&"/"&month(DateAdd("D",-2,now))&"/"&day(DateAdd("D",-2,now))&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')" &_
		" and Typeid=352 order by ActionMemberID"
	set rsDel1=conn.execute(strDel1)
	If Not rsDel1.Bof Then rsDel1.MoveFirst 
	While Not rsDel1.Eof
		if inStr(trim(rsDel1("ActionContent")),"單號:G")>0 and right(trim(rsDel1("ActionContent")),2)<>",," then
			CaseSn=CaseSn+1
			response.write CaseSn&". 刪除人:"&trim(rsDel1("ActionChName"))&"  "&replace(trim(rsDel1("ActionContent")),"舉發單刪除","")&"<br>"
		end if
		rsDel1.MoveNext
	Wend
	rsDel1.close
	set rsDel1=nothing
%>
		</span></td>
	</tr>

	<tr bgcolor="#FFFFCC">
		<td><%'4
		response.write(year(DateAdd("D",-3,now))-1911&"/"&month(DateAdd("D",-3,now))&"/"&day(DateAdd("D",-3,now)))
		%></td>
	</tr>
	<tr>
		<td><span class="style1">
<%	CaseSn=0
	strDel1="select * from Log where ActionDate between " &_
		"TO_DATE('"&year(DateAdd("D",-3,now))&"/"&month(DateAdd("D",-3,now))&"/"&day(DateAdd("D",-3,now))&" 0:0:0','YYYY/MM/DD/HH24/MI/SS')" &_
		" and TO_DATE('"&year(DateAdd("D",-3,now))&"/"&month(DateAdd("D",-3,now))&"/"&day(DateAdd("D",-3,now))&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')" &_
		" and Typeid=352 order by ActionMemberID"
	set rsDel1=conn.execute(strDel1)
	If Not rsDel1.Bof Then rsDel1.MoveFirst 
	While Not rsDel1.Eof
		if inStr(trim(rsDel1("ActionContent")),"單號:G")>0 and right(trim(rsDel1("ActionContent")),2)<>",," then
			CaseSn=CaseSn+1
			response.write CaseSn&". 刪除人:"&trim(rsDel1("ActionChName"))&"  "&replace(trim(rsDel1("ActionContent")),"舉發單刪除","")&"<br>"
		end if
		rsDel1.MoveNext
	Wend
	rsDel1.close
	set rsDel1=nothing
%>
		</span></td>
	</tr>

	<tr bgcolor="#FFFFCC">
		<td><%'5
		response.write(year(DateAdd("D",-4,now))-1911&"/"&month(DateAdd("D",-4,now))&"/"&day(DateAdd("D",-4,now)))
		%></td>
	</tr>
	<tr>
		<td><span class="style1">
<%	CaseSn=0
	strDel1="select * from Log where ActionDate between " &_
		"TO_DATE('"&year(DateAdd("D",-4,now))&"/"&month(DateAdd("D",-4,now))&"/"&day(DateAdd("D",-4,now))&" 0:0:0','YYYY/MM/DD/HH24/MI/SS')" &_
		" and TO_DATE('"&year(DateAdd("D",-4,now))&"/"&month(DateAdd("D",-4,now))&"/"&day(DateAdd("D",-4,now))&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')" &_
		" and Typeid=352 order by ActionMemberID"
	set rsDel1=conn.execute(strDel1)
	If Not rsDel1.Bof Then rsDel1.MoveFirst 
	While Not rsDel1.Eof
		if inStr(trim(rsDel1("ActionContent")),"單號:G")>0 and right(trim(rsDel1("ActionContent")),2)<>",," then
			CaseSn=CaseSn+1
			response.write CaseSn&". 刪除人:"&trim(rsDel1("ActionChName"))&"  "&replace(trim(rsDel1("ActionContent")),"舉發單刪除","")&"<br>"
		end if
		rsDel1.MoveNext
	Wend
	rsDel1.close
	set rsDel1=nothing
%>
		</span></td>
	</tr>

	<tr bgcolor="#FFFFCC">
		<td><%'6
		response.write(year(DateAdd("D",-5,now))-1911&"/"&month(DateAdd("D",-5,now))&"/"&day(DateAdd("D",-5,now)))
		%></td>
	</tr>
	<tr>
		<td><span class="style1">
<%	CaseSn=0
	strDel1="select * from Log where ActionDate between " &_
		"TO_DATE('"&year(DateAdd("D",-5,now))&"/"&month(DateAdd("D",-5,now))&"/"&day(DateAdd("D",-5,now))&" 0:0:0','YYYY/MM/DD/HH24/MI/SS')" &_
		" and TO_DATE('"&year(DateAdd("D",-5,now))&"/"&month(DateAdd("D",-5,now))&"/"&day(DateAdd("D",-5,now))&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')" &_
		" and Typeid=352 order by ActionMemberID"
	set rsDel1=conn.execute(strDel1)
	If Not rsDel1.Bof Then rsDel1.MoveFirst 
	While Not rsDel1.Eof
		if inStr(trim(rsDel1("ActionContent")),"單號:G")>0 and right(trim(rsDel1("ActionContent")),2)<>",," then
			CaseSn=CaseSn+1
			response.write CaseSn&". 刪除人:"&trim(rsDel1("ActionChName"))&"  "&replace(trim(rsDel1("ActionContent")),"舉發單刪除","")&"<br>"
		end if
		rsDel1.MoveNext
	Wend
	rsDel1.close
	set rsDel1=nothing
%>
		</span></td>
	</tr>

	<tr bgcolor="#FFFFCC">
		<td><%'7
		response.write(year(DateAdd("D",-6,now))-1911&"/"&month(DateAdd("D",-6,now))&"/"&day(DateAdd("D",-6,now)))
		%></td>
	</tr>
	<tr>
		<td><span class="style1">
<%	CaseSn=0
	strDel1="select * from Log where ActionDate between " &_
		"TO_DATE('"&year(DateAdd("D",-6,now))&"/"&month(DateAdd("D",-6,now))&"/"&day(DateAdd("D",-6,now))&" 0:0:0','YYYY/MM/DD/HH24/MI/SS')" &_
		" and TO_DATE('"&year(DateAdd("D",-6,now))&"/"&month(DateAdd("D",-6,now))&"/"&day(DateAdd("D",-6,now))&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')" &_
		" and Typeid=352 order by ActionMemberID"
	set rsDel1=conn.execute(strDel1)
	If Not rsDel1.Bof Then rsDel1.MoveFirst 
	While Not rsDel1.Eof
		if inStr(trim(rsDel1("ActionContent")),"單號:G")>0 and right(trim(rsDel1("ActionContent")),2)<>",," then
			CaseSn=CaseSn+1
			response.write CaseSn&". 刪除人:"&trim(rsDel1("ActionChName"))&"  "&replace(trim(rsDel1("ActionContent")),"舉發單刪除","")&"<br>"
		end if
		rsDel1.MoveNext
	Wend
	rsDel1.close
	set rsDel1=nothing
%>
		</span></td>
	</tr>
<%
	conn.close
	set conn=nothing
%>	
</table>
</body>
<script type="text/javascript" src="./js/date.js"></script>
<script language="JavaScript">

</script>
</html>
