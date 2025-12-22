<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fname="station.txt"
Response.AddHeader "Content-Disposition","attachment;filename=" &fname
Response.ContentType = "application/vnd.ms-txt"

strQuery="select distinct stationid,dcistationname from station order by stationid"
set rsfound=conn.execute(strQuery)
While Not rsfound.Eof
	response.write rsfound("stationid")&","&rsfound("dcistationname")&vbnewline
	rsfound.MoveNext
Wend
rsfound.close
set rsfound=nothing
conn.close
set conn=nothing
%>