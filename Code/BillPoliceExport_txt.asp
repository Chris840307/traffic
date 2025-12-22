<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fname="police.txt"
Response.AddHeader "Content-Disposition","attachment;filename=" &fname
Response.ContentType = "application/vnd.ms-txt"

strQuery="select a.CreditID,a.loginid,a.chname,b.UnitName from memberdata a,UnitInfo b where a.recordstateid=0 and a.accountstateid=0 and a.UnitID=b.UnitID order by a.chname"
set rsfound=conn.execute(strQuery)
While Not rsfound.Eof
	response.write rsfound("CreditID")&","&rsfound("loginid")&","&rsfound("chname")&","&rsfound("UnitName")&vbnewline
	rsfound.MoveNext
Wend
rsfound.close
set rsfound=nothing
conn.close
set conn=nothing
%>