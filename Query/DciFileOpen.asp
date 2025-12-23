<%
Function BytesToBstr(body,Cset)
dim objstream
set objstream = Server.CreateObject("adodb.stream")
objstream.Type = 1
objstream.Mode =3
objstream.Open
objstream.Write body
objstream.Position = 0
objstream.Type = 2
objstream.Charset = Cset
objstream.SaveTofile Server.mappath(".\tmp.txt"),2
objstream.close

Dim Fso,f
Set Fso=CreateObject("Scripting.FileSystemObject")
Set f=Fso.OpenTextFile(Server.mappath(".\tmp.txt"),1,True,0)
response.write "<table border=0>"
While Not f.AtEndOfStream
	response.write "<tr>"
	response.write "<TD nowrap>"
	response.write f.ReadLine
	response.write "</TD>"
	response.write "</tr>"
	
wend
response.write "</table>"
End function
%>
