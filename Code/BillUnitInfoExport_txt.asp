<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fname="unitinfo.txt"
Response.AddHeader "Content-Disposition","attachment;filename=" &fname
Response.ContentType = "application/vnd.ms-txt"

strQuery="select unitid,unitname from unitinfo order by unitid"
set rsfound=conn.execute(strQuery)
While Not rsfound.Eof
	response.write rsfound("unitid")&","&rsfound("unitname")&vbnewline
	rsfound.MoveNext
Wend
rsfound.close
set rsfound=nothing
conn.close
set conn=nothing
%>