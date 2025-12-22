<% 
Set fso = Server.CreateObject("Scripting.FileSystemObject") 
mProgressID=Request.QueryString("progressID") 
filePath=fso.GetSpecialFolder(2)&"\upload_"&mProgressID&".xml" 
Set f = fso.OpenTextFile(filePath,1) 
response.write f.ReadAll 
set f=nothing 
set fso=nothing 
%> 