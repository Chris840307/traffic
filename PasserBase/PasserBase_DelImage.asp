<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
	set fso=Server.CreateObject("Scripting.FileSystemObject")

	fp=server.mappath(Request("strPath"))&"\"&Request("strFileName")

	fso.DeleteFile fp, True

	strSQL="delete PasserImage where Imagefilename='"&Request("strFileName")&"'"
	conn.execute(strSQL)

	Response.Write "location.reload();"
%>