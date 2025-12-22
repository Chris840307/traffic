<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fname="color.txt"
Response.AddHeader "Content-Disposition","attachment;filename=" &fname
Response.ContentType = "application/vnd.ms-txt"

strQuery="select id,content from DCICode where TypeID=4 order by id"
set rsfound=conn.execute(strQuery)
While Not rsfound.Eof
	response.write rsfound("id")&","&rsfound("content")&vbnewline
	rsfound.MoveNext
Wend
rsfound.close
set rsfound=nothing
conn.close
set conn=nothing
%>