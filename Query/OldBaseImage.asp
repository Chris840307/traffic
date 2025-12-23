<!--#include virtual="Traffic/Common/Oldsp.ini"-->
<%
Response.ContentType = "image/jpeg"
sql="select "&Request("filedName")&" as img_1 from FMasterimg where FSeq='"&Request("FSeq")&"' and "&Request("filedName")&" is not null"

set rsBack=conn1.execute(sql)

while not rsBack.eof
	response.BinaryWrite rsBack("img_1")
	rsBack.movenext
wend
rsBack.close
conn1.close
%>