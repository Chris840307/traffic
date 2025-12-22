<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fname="rule.txt"
Response.AddHeader "Content-Disposition","attachment;filename=" &fname
Response.ContentType = "application/vnd.ms-txt"

strQuery="select itemid,illegalrule from law where version=2 order by itemid"
set rsfound=conn.execute(strQuery)
While Not rsfound.Eof
	response.write rsfound("itemid")&","&rsfound("illegalrule")&vbnewline
	rsfound.MoveNext
Wend
rsfound.close
set rsfound=nothing
conn.close
set conn=nothing
%>